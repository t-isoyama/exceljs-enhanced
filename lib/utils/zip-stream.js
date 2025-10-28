const events = require('events');
const JSZip = require('jszip');
// Performance Phase 4-2: Use archiver for Node.js (better streaming), JSZip for browser
let Archiver;
try {
  // Only load archiver in Node.js environment (not in browser builds)
  if (!process.browser) {
    Archiver = require('archiver');
  }
} catch (e) {
  // Fallback to JSZip if archiver is not available
  Archiver = null;
}

const StreamBuf = require('./stream-buf');
const {stringToBuffer} = require('./browser-buffer-encode');

// =============================================================================
// The ZipWriter class
// Packs streamed data into an output zip stream
class ZipWriter extends events.EventEmitter {
  constructor(options) {
    super();
    this.options = Object.assign(
      {
        type: 'nodebuffer',
        compression: 'DEFLATE',
      },
      options
    );

    // Performance Phase 4-2: Use archiver for Node.js, JSZip for browser
    this.useArchiver = !process.browser && Archiver;

    if (this.useArchiver) {
      // Archiver path (Node.js) - streaming write
      this.zip = Archiver('zip', options);
      this.stream = new StreamBuf();
      this.zip.pipe(this.stream);
    } else {
      // JSZip path (browser or fallback) - buffer-based write
      this.zip = new JSZip();
      this.stream = new StreamBuf();
    }
  }

  append(data, options) {
    if (this.useArchiver) {
      // Archiver API
      if (options.hasOwnProperty('base64') && options.base64) {
        // archiver doesn't have native base64 support, decode first
        const buffer = Buffer.from(data, 'base64');
        this.zip.append(buffer, {name: options.name});
      } else {
        this.zip.append(data, {name: options.name});
      }
      return;
    }

    // JSZip API
    if (options.hasOwnProperty('base64') && options.base64) {
      this.zip.file(options.name, data, {base64: true});
      return;
    }

    // https://www.npmjs.com/package/process
    if (process.browser && typeof data === 'string') {
      // use TextEncoder in browser
      data = stringToBuffer(data);
    }
    this.zip.file(options.name, data);
  }

  async finalize() {
    if (this.useArchiver) {
      // Archiver finalization
      return new Promise((resolve, reject) => {
        this.zip.on('finish', () => {
          this.emit('finish');
          resolve();
        });
        this.zip.on('error', reject);
        this.zip.finalize();
      });
    }

    // JSZip finalization
    const content = await this.zip.generateAsync(this.options);
    this.stream.end(content);
    this.emit('finish');
    return undefined;
  }

  // ==========================================================================
  // Stream.Readable interface
  read(size) {
    return this.stream.read(size);
  }

  setEncoding(encoding) {
    return this.stream.setEncoding(encoding);
  }

  pause() {
    return this.stream.pause();
  }

  resume() {
    return this.stream.resume();
  }

  isPaused() {
    return this.stream.isPaused();
  }

  pipe(destination, options) {
    return this.stream.pipe(destination, options);
  }

  unpipe(destination) {
    return this.stream.unpipe(destination);
  }

  unshift(chunk) {
    return this.stream.unshift(chunk);
  }

  wrap(stream) {
    return this.stream.wrap(stream);
  }
}

// =============================================================================

module.exports = {
  ZipWriter,
};
