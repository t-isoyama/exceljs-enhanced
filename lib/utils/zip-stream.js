const {EventEmitter} = require('events');
const {Zip, ZipPassThrough, strToU8} = require('fflate');

const StreamBuf = require('./stream-buf');

// =============================================================================
// The ZipWriter class
// Unified ZIP writer using fflate for both Node.js and browser
// Performance Phase 2-2: Migrate from archiver/jszip to fflate
class ZipWriter extends EventEmitter {
  constructor(options) {
    super();

    // Modern API: Use Object.hasOwn instead of hasOwnProperty
    this.options = {
      level: Object.hasOwn(options || {}, 'level') ? options.level : 6,
      ...options,
    };

    this.stream = new StreamBuf();
    this.finalized = false;
    this.pendingFiles = new Set();

    // Create fflate Zip instance with streaming callback
    this.zip = new Zip((err, data, final) => {
      if (err) {
        this.emit('error', err);
        return;
      }

      // Write chunk to stream
      if (data && data.length > 0) {
        this.stream.write(Buffer.from(data));
      }

      // Finalize stream when ZIP is complete
      if (final) {
        this.stream.end();
        this.emit('finish');
      }
    });
  }

  append(data, options) {
    if (this.finalized) {
      throw new Error('Cannot append to finalized ZIP');
    }

    const {name, base64} = options || {};
    if (!name) {
      throw new Error('File name is required');
    }

    // Handle stream input (like archiver does)
    if (data && typeof data.on === 'function' && typeof data.read === 'function') {
      this._appendStream(data, options);
      return undefined;
    }

    // Track pending file
    this.pendingFiles.add(name);

    // Create ZIP entry with compression level
    const file = new ZipPassThrough(name);
    file.level = this.options.level;

    // Add file to ZIP
    this.zip.add(file);

    // Process data
    let buffer;
    if (base64) {
      // Decode base64 data
      buffer = Buffer.from(data, 'base64');
    } else if (typeof data === 'string') {
      // Convert string to Uint8Array
      buffer = Buffer.from(strToU8(data));
    } else if (Buffer.isBuffer(data)) {
      buffer = data;
    } else if (data instanceof Uint8Array) {
      buffer = Buffer.from(data);
    } else {
      // For other types, convert to buffer
      buffer = Buffer.from(data);
    }

    // Push data to file stream and mark as complete
    file.push(buffer, true);

    // Remove from pending
    this.pendingFiles.delete(name);
    return undefined;
  }

  _appendStream(stream, options) {
    const {name} = options || {};

    // Track pending file
    this.pendingFiles.add(name);

    // Create ZIP entry
    const file = new ZipPassThrough(name);
    file.level = this.options.level;

    // Add file to ZIP
    this.zip.add(file);

    // Collect all data from stream, then push to ZIP
    // This approach works better with StreamBuf that may not be fully initialized
    const chunks = [];

    const onData = chunk => {
      chunks.push(chunk);
    };

    const onEnd = () => {
      // Concatenate all chunks and push to ZIP
      if (chunks.length > 0) {
        const buffer = Buffer.concat(chunks);
        file.push(buffer);
      }
      // Mark as complete
      file.push(new Uint8Array(0), true);
      this.pendingFiles.delete(name);
    };

    const onError = err => {
      this.emit('error', err);
      this.pendingFiles.delete(name);
    };

    // Use setImmediate to ensure stream is fully initialized
    setImmediate(() => {
      try {
        if (stream && typeof stream.on === 'function') {
          stream.on('data', onData);
          stream.on('end', onEnd);
          stream.on('error', onError);
        } else {
          // Stream is not valid, finalize immediately
          file.push(new Uint8Array(0), true);
          this.pendingFiles.delete(name);
        }
      } catch (err) {
        // Stream setup failed, finalize immediately
        file.push(new Uint8Array(0), true);
        this.pendingFiles.delete(name);
      }
    });
  }

  // Support for archiver-style file() method
  file(filepath, options) {
    const fs = require('fs');
    const stream = fs.createReadStream(filepath);
    return this.append(stream, options);
  }

  async finalize() {
    if (this.finalized) {
      return undefined;
    }

    // Wait for all pending files to complete
    if (this.pendingFiles.size > 0) {
      await new Promise(resolve => {
        setTimeout(resolve, 10);
      });
      if (this.pendingFiles.size > 0) {
        throw new Error(`Cannot finalize: ${this.pendingFiles.size} files still pending`);
      }
    }

    this.finalized = true;

    // End the ZIP stream
    this.zip.end();

    // Return a promise that resolves when finished
    return new Promise((resolve, reject) => {
      this.once('finish', resolve);
      this.once('error', reject);
    });
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
