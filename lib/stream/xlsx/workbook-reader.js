const fs = require('fs');
const {EventEmitter} = require('events');
// Performance: Use native stream instead of readable-stream polyfill (Node 22+)
const {Readable, PassThrough} = require('stream');
const {unzipSync} = require('fflate');
const iterateStream = require('../../utils/iterate-stream');
const parseSax = require('../../utils/parse-sax');

const StyleManager = require('../../xlsx/xform/style/styles-xform');
const WorkbookXform = require('../../xlsx/xform/book/workbook-xform');
const RelationshipsXform = require('../../xlsx/xform/core/relationships-xform');

const WorksheetReader = require('./worksheet-reader');
const HyperlinkReader = require('./hyperlink-reader');

class WorkbookReader extends EventEmitter {
  constructor(input, options = {}) {
    super();

    this.input = input;

    this.options = {
      worksheets: 'emit',
      sharedStrings: 'cache',
      hyperlinks: 'ignore',
      styles: 'ignore',
      entries: 'ignore',
      ...options,
    };

    this.styles = new StyleManager();
    this.styles.init();
  }

  _getStream(input) {
    if (input instanceof Readable) {
      return input;
    }
    if (typeof input === 'string') {
      return fs.createReadStream(input);
    }
    throw new Error(`Could not recognise input: ${input}`);
  }

  _createStreamFromBuffer(buffer) {
    const stream = new PassThrough();
    stream.end(Buffer.from(buffer));
    return stream;
  }

  async read(input, options) {
    try {
      for await (const {eventType, value} of this.parse(input, options)) {
        switch (eventType) {
          case 'shared-strings':
            this.emit(eventType, value);
            break;
          case 'worksheet':
            this.emit(eventType, value);
            await value.read();
            break;
          case 'hyperlinks':
            this.emit(eventType, value);
            break;
        }
      }
      this.emit('end');
      this.emit('finished');
    } catch (error) {
      this.emit('error', error);
    }
  }

  async *[Symbol.asyncIterator]() {
    for await (const {eventType, value} of this.parse()) {
      if (eventType === 'worksheet') {
        yield value;
      }
    }
  }

  async *parse(input, options) {
    if (options) this.options = options;
    const stream = (this.stream = this._getStream(input || this.input));

    // Performance Phase 2-2: Read entire stream into buffer and use fflate unzipSync
    const chunks = [];
    for await (const chunk of stream) {
      chunks.push(chunk);
    }
    const buffer = Buffer.concat(chunks);
    const files = unzipSync(buffer);

    // worksheets, deferred for parsing after shared strings reading
    const waitingWorkSheets = [];

    // Process files in order
    const getOrder = filename => {
      if (filename.includes('sharedStrings')) return 0;
      if (filename.includes('styles')) return 1;
      return 2;
    };

    const sortedEntries = Object.entries(files).sort(([a], [b]) => {
      const orderA = getOrder(a);
      const orderB = getOrder(b);
      return orderA - orderB || a.localeCompare(b);
    });

    /* eslint-disable no-await-in-loop */
    for (const [entryPath, fileData] of sortedEntries) {
      let match;
      let sheetNo;
      switch (entryPath) {
        case '_rels/.rels':
          break;
        case 'xl/_rels/workbook.xml.rels':
          await this._parseRels(this._createStreamFromBuffer(fileData));
          break;
        case 'xl/workbook.xml':
          await this._parseWorkbook(this._createStreamFromBuffer(fileData));
          break;
        case 'xl/sharedStrings.xml':
          await this._parseSharedStrings(this._createStreamFromBuffer(fileData));
          break;
        case 'xl/styles.xml':
          await this._parseStyles(this._createStreamFromBuffer(fileData));
          break;
        default:
          if (entryPath.match(/xl\/worksheets\/sheet\d+[.]xml/)) {
            match = entryPath.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
            sheetNo = match[1];

            if (this.sharedStrings && this.workbookRels) {
              yield* this._parseWorksheet(
                iterateStream(this._createStreamFromBuffer(fileData)),
                sheetNo
              );
            } else {
              // Defer worksheet parsing until sharedStrings and workbookRels are loaded
              waitingWorkSheets.push({sheetNo, data: fileData});
            }
          } else if (entryPath.match(/xl\/worksheets\/_rels\/sheet\d+[.]xml.rels/)) {
            match = entryPath.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
            sheetNo = match[1];
            yield* this._parseHyperlinks(
              iterateStream(this._createStreamFromBuffer(fileData)),
              sheetNo
            );
          }
          break;
      }
    }
    /* eslint-enable no-await-in-loop */

    // Process all deferred worksheets after metadata is loaded
    for (const {sheetNo, data} of waitingWorkSheets) {
      yield* this._parseWorksheet(iterateStream(this._createStreamFromBuffer(data)), sheetNo);
    }
  }

  _emitEntry(payload) {
    if (this.options.entries === 'emit') {
      this.emit('entry', payload);
    }
  }

  async _parseRels(entry) {
    const xform = new RelationshipsXform();
    this.workbookRels = await xform.parseStream(iterateStream(entry));
  }

  async _parseWorkbook(entry) {
    this._emitEntry({type: 'workbook'});

    const workbook = new WorkbookXform();
    await workbook.parseStream(iterateStream(entry));

    this.properties = workbook.map.workbookPr;
    this.model = workbook.model;
  }

  async _parseSharedStrings(entry) {
    this._emitEntry({type: 'shared-strings'});

    if (this.options.sharedStrings === 'ignore') {
      return;
    }

    if (this.options.sharedStrings === 'cache') {
      this.sharedStrings = [];
    }

    let text = null;
    let richText = [];
    let font = null;
    for await (const events of parseSax(iterateStream(entry))) {
      for (const {eventType, value} of events) {
        if (eventType === 'opentag') {
          const node = value;
          switch (node.name) {
            case 'b':
              font = font || {};
              font.bold = true;
              break;
            case 'charset':
              font = font || {};
              font.charset = parseInt(node.attributes.charset, 10);
              break;
            case 'color':
              font = font || {};
              font.color = {};
              if (node.attributes.rgb) {
                font.color.argb = node.attributes.argb;
              }
              if (node.attributes.val) {
                font.color.argb = node.attributes.val;
              }
              if (node.attributes.theme) {
                font.color.theme = node.attributes.theme;
              }
              break;
            case 'family':
              font = font || {};
              font.family = parseInt(node.attributes.val, 10);
              break;
            case 'i':
              font = font || {};
              font.italic = true;
              break;
            case 'outline':
              font = font || {};
              font.outline = true;
              break;
            case 'rFont':
              font = font || {};
              font.name = node.value;
              break;
            case 'si':
              font = null;
              richText = [];
              text = null;
              break;
            case 'sz':
              font = font || {};
              font.size = parseInt(node.attributes.val, 10);
              break;
            case 'strike':
              break;
            case 't':
              text = null;
              break;
            case 'u':
              font = font || {};
              font.underline = true;
              break;
            case 'vertAlign':
              font = font || {};
              font.vertAlign = node.attributes.val;
              break;
          }
        } else if (eventType === 'text') {
          text = text ? text + value : value;
        } else if (eventType === 'closetag') {
          const node = value;
          switch (node.name) {
            case 'r':
              richText.push({
                font,
                text,
              });

              font = null;
              text = null;
              break;
            case 'si': {
              const sharedStringValue = richText.length ? {richText} : text;
              if (this.options.sharedStrings === 'cache') {
                this.sharedStrings.push(sharedStringValue);
              } else if (this.options.sharedStrings === 'emit') {
                this.emit('shared-strings', sharedStringValue);
              }

              richText = [];
              font = null;
              text = null;
              break;
            }
          }
        }
      }
    }
  }

  async _parseStyles(entry) {
    this._emitEntry({type: 'styles'});
    if (this.options.styles === 'cache') {
      this.styles = new StyleManager();
      await this.styles.parseStream(iterateStream(entry));
    }
  }

  *_parseWorksheet(iterator, sheetNo) {
    this._emitEntry({type: 'worksheet', id: sheetNo});
    const worksheetReader = new WorksheetReader({
      workbook: this,
      id: sheetNo,
      iterator,
      options: this.options,
    });

    const matchingRel = (this.workbookRels || []).find(rel => rel.Target === `worksheets/sheet${sheetNo}.xml`);
    const matchingSheet = matchingRel && (this.model.sheets || []).find(sheet => sheet.rId === matchingRel.Id);
    if (matchingSheet) {
      worksheetReader.id = matchingSheet.id;
      worksheetReader.name = matchingSheet.name;
      worksheetReader.state = matchingSheet.state;
    }
    if (this.options.worksheets === 'emit') {
      yield {eventType: 'worksheet', value: worksheetReader};
    }
  }

  *_parseHyperlinks(iterator, sheetNo) {
    this._emitEntry({type: 'hyperlinks', id: sheetNo});
    const hyperlinksReader = new HyperlinkReader({
      workbook: this,
      id: sheetNo,
      iterator,
      options: this.options,
    });
    if (this.options.hyperlinks === 'emit') {
      yield {eventType: 'hyperlinks', value: hyperlinksReader};
    }
  }
}

// for reference - these are the valid values for options
WorkbookReader.Options = {
  worksheets: ['emit', 'ignore'],
  sharedStrings: ['cache', 'emit', 'ignore'],
  hyperlinks: ['cache', 'emit', 'ignore'],
  styles: ['cache', 'ignore'],
  entries: ['emit', 'ignore'],
};

module.exports = WorkbookReader;
