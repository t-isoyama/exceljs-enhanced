const XmlStream = require('../../../utils/xml-stream');
const BaseXform = require('../base-xform');
const SharedStringXform = require('./shared-string-xform');

class SharedStringsXform extends BaseXform {
  constructor(model) {
    super();

    this.model = model || {
      values: [],
      count: 0,
    };
    this.hash = new Map();
    this.rich = new Map();
  }

  get sharedStringXform() {
    return this._sharedStringXform || (this._sharedStringXform = new SharedStringXform());
  }

  get values() {
    return this.model.values;
  }

  get uniqueCount() {
    return this.model.values.length;
  }

  get count() {
    return this.model.count;
  }

  getString(index) {
    return this.model.values[index];
  }

  add(value) {
    return value.richText ? this.addRichText(value) : this.addText(value);
  }

  addText(value) {
    let index = this.hash.get(value);
    if (index === undefined) {
      index = this.model.values.length;
      this.hash.set(value, index);
      this.model.values.push(value);
    }
    this.model.count++;
    return index;
  }

  addRichText(value) {
    // Future: Use WeakMap for rich text deduplication optimization
    const xml = this.sharedStringXform.toXml(value);
    let index = this.rich.get(xml);
    if (index === undefined) {
      index = this.model.values.length;
      this.rich.set(xml, index);
      this.model.values.push(value);
    }
    this.model.count++;
    return index;
  }

  // <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  // <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="<%=totalRefs%>" uniqueCount="<%=count%>">
  //   <si><t><%=text%></t></si>
  //   <si><r><rPr></rPr><t></t></r></si>
  // </sst>

  render(xmlStream, model) {
    model = model || this._values;
    xmlStream.openXml(XmlStream.StdDocAttributes);

    xmlStream.openNode('sst', {
      xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      count: model.count,
      uniqueCount: model.values.length,
    });

    const sx = this.sharedStringXform;
    model.values.forEach(sharedString => {
      sx.render(xmlStream, sharedString);
    });
    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'sst':
        return true;
      case 'si':
        this.parser = this.sharedStringXform;
        this.parser.parseOpen(node);
        return true;
      default:
        throw new Error(`Unexpected xml node in parseOpen: ${JSON.stringify(node)}`);
    }
  }

  parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model.values.push(this.parser.model);
        this.model.count++;
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case 'sst':
        return false;
      default:
        throw new Error(`Unexpected xml node in parseClose: ${name}`);
    }
  }
}

module.exports = SharedStringsXform;
