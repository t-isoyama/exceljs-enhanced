// Modern browser bundle - no polyfills needed for ES2015+ features
// Target browsers: Chrome 90+, Firefox 88+, Safari 14+ (2021+)
// For older browsers, please load polyfills separately (e.g., polyfill.io or core-js)

const ExcelJS = {
  Workbook: require('./doc/workbook'),
};

// Object.assign mono-fill
const Enums = require('./doc/enums');

Object.keys(Enums).forEach(key => {
  ExcelJS[key] = Enums[key];
});

module.exports = ExcelJS;
