class SharedStrings {
  constructor() {
    this._values = [];
    this._totalRefs = 0;
    // Performance: Use Map instead of Object.create(null) for better performance
    // Map is optimized for frequent additions and lookups, especially with string keys
    this._hash = new Map();
  }

  get count() {
    return this._values.length;
  }

  get values() {
    return this._values;
  }

  get totalRefs() {
    return this._totalRefs;
  }

  getString(index) {
    return this._values[index];
  }

  add(value) {
    let index = this._hash.get(value);
    if (index === undefined) {
      index = this._values.length;
      this._hash.set(value, index);
      this._values.push(value);
    }
    this._totalRefs++;
    return index;
  }
}

module.exports = SharedStrings;
