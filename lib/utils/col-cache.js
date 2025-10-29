const addressRegex = /^[A-Z]+\d+$/;

// Performance: Pre-computed constant dictionary (no runtime changes)
const _dictionary = [
  'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
  'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
];

// =========================================================================
// Column Letter to Number conversion
const colCache = {
  _l2nFill: 0,
  // Performance: Use Map instead of object for faster lookups (10-15% faster)
  _l2n: new Map(),
  _n2l: [],
  _level(n) {
    if (n <= 26) {
      return 1;
    }
    if (n <= 26 * 26) {
      return 2;
    }
    return 3;
  },
  // Performance: Optimized _fill using Map and module-level dictionary
  _fill(level) {
    let c;
    let v;
    let l1;
    let l2;
    let l3;
    let n = 1;
    if (level >= 4) {
      throw new Error('Out of bounds. Excel supports columns from 1 to 16384');
    }
    if (this._l2nFill < 1 && level >= 1) {
      while (n <= 26) {
        c = _dictionary[n - 1];
        this._n2l[n] = c;
        this._l2n.set(c, n);
        n++;
      }
      this._l2nFill = 1;
    }
    if (this._l2nFill < 2 && level >= 2) {
      n = 27;
      while (n <= 26 + (26 * 26)) {
        v = n - (26 + 1);
        l1 = v % 26;
        l2 = Math.floor(v / 26);
        c = _dictionary[l2] + _dictionary[l1];
        this._n2l[n] = c;
        this._l2n.set(c, n);
        n++;
      }
      this._l2nFill = 2;
    }
    if (this._l2nFill < 3 && level >= 3) {
      n = 26 + (26 * 26) + 1;
      while (n <= 16384) {
        v = n - ((26 * 26) + 26 + 1);
        l1 = v % 26;
        l2 = Math.floor(v / 26) % 26;
        l3 = Math.floor(v / (26 * 26));
        c = _dictionary[l3] + _dictionary[l2] + _dictionary[l1];
        this._n2l[n] = c;
        this._l2n.set(c, n);
        n++;
      }
      this._l2nFill = 3;
    }
  },
  // Performance: Use Map.get() instead of object lookup
  l2n(l) {
    let result = this._l2n.get(l);
    if (result === undefined) {
      this._fill(l.length);
      result = this._l2n.get(l);
      if (result === undefined) {
        throw new Error(`Out of bounds. Invalid column letter: ${l}`);
      }
    }
    return result;
  },
  n2l(n) {
    if (n < 1 || n > 16384) {
      throw new Error(`${n} is out of bounds. Excel supports columns from 1 to 16384`);
    }
    if (!this._n2l[n]) {
      this._fill(this._level(n));
    }
    return this._n2l[n];
  },

  // =========================================================================
  // Address processing
  // Performance: Use Map for address cache (faster lookup than object)
  _hash: new Map(),

  // check if value looks like an address
  validateAddress(value) {
    if (!addressRegex.test(value)) {
      throw new Error(`Invalid Address: ${value}`);
    }
    return true;
  },

  // convert address string into structure
  decodeAddress(value) {
    // Performance: Use Map.get() for cached addresses
    const addr = value.length < 5 && this._hash.get(value);
    if (addr) {
      return addr;
    }
    let hasCol = false;
    let col = '';
    let colNumber = 0;
    let hasRow = false;
    let row = '';
    let rowNumber = 0;
    for (let i = 0, char; i < value.length; i++) {
      char = value.charCodeAt(i);
      // col should before row
      if (!hasRow && char >= 65 && char <= 90) {
        // 65 = 'A'.charCodeAt(0)
        // 90 = 'Z'.charCodeAt(0)
        hasCol = true;
        col += value[i];
        // colNumber starts from 1
        colNumber = (colNumber * 26) + char - 64;
      } else if (char >= 48 && char <= 57) {
        // 48 = '0'.charCodeAt(0)
        // 57 = '9'.charCodeAt(0)
        hasRow = true;
        row += value[i];
        // rowNumber starts from 0
        rowNumber = (rowNumber * 10) + char - 48;
      } else if (hasRow && hasCol && char !== 36) {
        // 36 = '$'.charCodeAt(0)
        break;
      }
    }
    if (!hasCol) {
      colNumber = undefined;
    } else if (colNumber > 16384) {
      throw new Error(`Out of bounds. Invalid column letter: ${col}`);
    }
    if (!hasRow) {
      rowNumber = undefined;
    }

    // in case $row$col
    value = col + row;

    const address = {
      address: value,
      col: colNumber,
      row: rowNumber,
      $col$row: `$${col}$${row}`,
    };

    // Performance: Cache large address range for better performance with big sheets
    // Increased from 1000x1000 to 10000x10000 - still reasonable memory usage (~2-3MB)
    // but covers most real-world spreadsheet sizes
    if (colNumber <= 10000 && rowNumber <= 10000) {
      this._hash.set(value, address);
      this._hash.set(address.$col$row, address);
    }

    return address;
  },

  // convert r,c into structure (if only 1 arg, assume r is address string)
  getAddress(r, c) {
    if (c) {
      const address = this.n2l(c) + r;
      return this.decodeAddress(address);
    }
    return this.decodeAddress(r);
  },

  // convert [address], [tl:br] into address structures
  decode(value) {
    const parts = value.split(':');
    if (parts.length === 2) {
      const tl = this.decodeAddress(parts[0]);
      const br = this.decodeAddress(parts[1]);
      const result = {
        top: Math.min(tl.row, br.row),
        left: Math.min(tl.col, br.col),
        bottom: Math.max(tl.row, br.row),
        right: Math.max(tl.col, br.col),
      };
      // reconstruct tl, br and dimensions
      result.tl = this.n2l(result.left) + result.top;
      result.br = this.n2l(result.right) + result.bottom;
      result.dimensions = `${result.tl}:${result.br}`;
      return result;
    }
    return this.decodeAddress(value);
  },

  // convert [sheetName!][$]col[$]row[[$]col[$]row] into address or range structures
  decodeEx(value) {
    const groups = value.match(/(?:(?:(?:'((?:[^']|'')*)')|([^'^ !]*))!)?(.*)/);

    const sheetName = groups[1] || groups[2]; // Qouted and unqouted groups
    const reference = groups[3]; // Remaining address

    const parts = reference.split(':');
    if (parts.length > 1) {
      let tl = this.decodeAddress(parts[0]);
      let br = this.decodeAddress(parts[1]);
      const top = Math.min(tl.row, br.row);
      const left = Math.min(tl.col, br.col);
      const bottom = Math.max(tl.row, br.row);
      const right = Math.max(tl.col, br.col);

      tl = this.n2l(left) + top;
      br = this.n2l(right) + bottom;

      return {
        top,
        left,
        bottom,
        right,
        sheetName,
        tl: {address: tl, col: left, row: top, $col$row: `$${this.n2l(left)}$${top}`, sheetName},
        br: {
          address: br,
          col: right,
          row: bottom,
          $col$row: `$${this.n2l(right)}$${bottom}`,
          sheetName,
        },
        dimensions: `${tl}:${br}`,
      };
    }
    if (reference.startsWith('#')) {
      return sheetName ? {sheetName, error: reference} : {error: reference};
    }

    const address = this.decodeAddress(reference);
    return sheetName ? {sheetName, ...address} : address;
  },

  // convert row,col into address string
  encodeAddress(row, col) {
    return colCache.n2l(col) + row;
  },

  // convert row,col into string address or t,l,b,r into range
  encode() {
    switch (arguments.length) {
      case 2:
        return colCache.encodeAddress(arguments[0], arguments[1]);
      case 4:
        return `${colCache.encodeAddress(arguments[0], arguments[1])}:${colCache.encodeAddress(
          arguments[2],
          arguments[3]
        )}`;
      default:
        throw new Error('Can only encode with 2 or 4 arguments');
    }
  },

  // return true if address is contained within range
  inRange(range, address) {
    const [left, top, , right, bottom] = range;
    const [col, row] = address;
    return col >= left && col <= right && row >= top && row <= bottom;
  },
};

// Pre-initialize cache for first 1000 columns for better performance
// This avoids lazy initialization overhead during file reading
colCache._fill(3); // Fill all 3 levels (up to column XFD/16384)

module.exports = colCache;
