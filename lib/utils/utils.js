const fs = require('fs');

// useful stuff
const inherits = function(cls, superCtor, statics, prototype) {
  // eslint-disable-next-line no-underscore-dangle
  cls.super_ = superCtor;

  if (!prototype) {
    prototype = statics;
    statics = null;
  }

  if (statics) {
    Object.keys(statics).forEach(i => {
      Object.defineProperty(cls, i, Object.getOwnPropertyDescriptor(statics, i));
    });
  }

  const properties = {
    constructor: {
      value: cls,
      enumerable: false,
      writable: false,
      configurable: true,
    },
  };
  if (prototype) {
    Object.keys(prototype).forEach(i => {
      properties[i] = Object.getOwnPropertyDescriptor(prototype, i);
    });
  }

  cls.prototype = Object.create(superCtor.prototype, properties);
};

// eslint-disable-next-line no-control-regex
const xmlDecodeRegex = /[<>&'"\x7F\x00-\x08\x0B-\x0C\x0E-\x1F]/;

// Performance: Pre-computed lookup table for XML entity encoding
const xmlEncodeLookup = new Array(128);
xmlEncodeLookup[34] = '&quot;'; // "
xmlEncodeLookup[38] = '&amp;';  // &
xmlEncodeLookup[39] = '&apos;'; // '
xmlEncodeLookup[60] = '&lt;';   // <
xmlEncodeLookup[62] = '&gt;';   // >
xmlEncodeLookup[127] = '';      // DEL
// Control characters (except \t, \n, \r)
for (let i = 0; i <= 8; i++) xmlEncodeLookup[i] = '';
for (let i = 11; i <= 12; i++) xmlEncodeLookup[i] = '';
for (let i = 14; i <= 31; i++) xmlEncodeLookup[i] = '';

const utils = {
  nop() {},
  promiseImmediate(value) {
    return new Promise(resolve => {
      if (global.setImmediate) {
        setImmediate(() => {
          resolve(value);
        });
      } else {
        // poorman's setImmediate - must wait at least 1ms
        setTimeout(() => {
          resolve(value);
        }, 1);
      }
    });
  },
  inherits,
  dateToExcel(d, date1904) {
    // eslint-disable-next-line no-mixed-operators
    return 25569 + d.getTime() / (24 * 3600 * 1000) - (date1904 ? 1462 : 0);
  },
  excelToDate(v, date1904) {
    // eslint-disable-next-line no-mixed-operators
    const millisecondSinceEpoch = Math.round((v - 25569 + (date1904 ? 1462 : 0)) * 24 * 3600 * 1000);
    return new Date(millisecondSinceEpoch);
  },
  parsePath(filepath) {
    const last = filepath.lastIndexOf('/');
    return {
      path: filepath.substring(0, last),
      name: filepath.substring(last + 1),
    };
  },
  getRelsPath(filepath) {
    const path = utils.parsePath(filepath);
    return `${path.path}/_rels/${path.name}.rels`;
  },
  xmlEncode(text) {
    // Performance: Early return if no encoding needed
    const regexResult = xmlDecodeRegex.exec(text);
    if (!regexResult) return text;

    // Performance: Use array for accumulation, then join once
    const parts = [];
    let lastIndex = 0;
    let i = regexResult.index;

    for (; i < text.length; i++) {
      const charCode = text.charCodeAt(i);
      // Performance: Use lookup table instead of switch
      if (charCode < 128 && xmlEncodeLookup[charCode] !== undefined) {
        if (lastIndex !== i) parts.push(text.substring(lastIndex, i));
        const escape = xmlEncodeLookup[charCode];
        if (escape) parts.push(escape);
        lastIndex = i + 1;
      }
    }

    if (lastIndex === 0) return text; // Nothing to encode
    if (lastIndex !== i) parts.push(text.substring(lastIndex, i));
    return parts.join('');
  },
  xmlDecode(text) {
    return text.replace(/&([a-z]*);/g, c => {
      switch (c) {
        case '&lt;':
          return '<';
        case '&gt;':
          return '>';
        case '&amp;':
          return '&';
        case '&apos;':
          return '\'';
        case '&quot;':
          return '"';
        default:
          return c;
      }
    });
  },
  validInt(value) {
    const i = parseInt(value, 10);
    return !Number.isNaN(i) ? i : 0;
  },

  isDateFmt(fmt) {
    if (!fmt) {
      return false;
    }

    // must remove all chars inside quotes and []
    fmt = fmt.replace(/\[[^\]]*]/g, '');
    fmt = fmt.replace(/"[^"]*"/g, '');
    // then check for date formatting chars
    const result = fmt.match(/[ymdhMsb]+/) !== null;
    return result;
  },

  fs: {
    exists(path) {
      return new Promise(resolve => {
        fs.access(path, fs.constants.F_OK, err => {
          resolve(!err);
        });
      });
    },
  },

  toIsoDateString(dt) {
    return dt.toIsoString().subsstr(0, 10);
  },

  parseBoolean(value) {
    return value === true || value === 'true' || value === 1 || value === '1';
  },

  *range(start, stop, step = 1) {
    const compareOrder = step > 0 ? (a, b) => a < b : (a, b) => a > b;
    for (let value = start; compareOrder(value, stop); value += step) {
      yield value;
    }
  },

  toSortedArray(values) {
    const result = Array.from(values);

    // Note: per default, `Array.prototype.sort()` converts values
    // to strings when comparing. Here, if we have numbers, we use
    // numeric sort.
    if (result.every(item => Number.isFinite(item))) {
      const compareNumbers = (a, b) => a - b;
      return result.sort(compareNumbers);
    }

    return result.sort();
  },

  objectFromProps(props, value = null) {
    // *Note*: Using `reduce` as `Object.fromEntries` requires Node 12+;
    // ExcelJs is >=8.3.0 (as of 2023-10-08).
    // return Object.fromEntries(props.map(property => [property, value]));
    return props.reduce((result, property) => {
      result[property] = value;
      return result;
    }, {});
  },
};

module.exports = utils;
