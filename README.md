# ExcelJS Enhanced

> A modern, performance-optimized fork of [ExcelJS](https://github.com/exceljs/exceljs)

**ExcelJS Enhanced** is a performance-focused fork of the popular ExcelJS library, providing comprehensive Excel Workbook management for Node.js and browsers with significant speed improvements and modern API updates.

## 🔄 This is a Fork

This project is forked from [exceljs/exceljs](https://github.com/exceljs/exceljs) v4.4.0 and includes significant performance optimizations and modernization improvements while maintaining full API compatibility with the original library.

For detailed API documentation, please refer to the [original ExcelJS documentation](https://github.com/exceljs/exceljs).

## ✨ Key Enhancements

This fork includes the following major improvements:

- **🚀 Performance**: 7-18x cumulative speedup through optimized cell/XML/address operations
- **📦 Modern Dependencies**: Migrated to modern libraries (fflate, saxes, dayjs, vitest)
- **🗜️ Smaller Bundle**: Optimized browser bundle size (removed unnecessary polyfills)
- **⚡ Modern APIs**: Adopted Node.js 22+ native APIs (Object.hasOwn, native streams)
- **🛠️ Modern Build**: Replaced Grunt+Babel with Browserify+esbuild for faster builds
- **✅ Modern Testing**: Migrated from Mocha to Vitest for better developer experience
- **🔐 Security**: Updated dependencies to fix security vulnerabilities

## 📋 Requirements

- **Node.js**: >=22.0.0

## 📦 Installation

```bash
npm install @t-isoyama/exceljs-enhanced
```

## 🚀 Quick Start

```javascript
const ExcelJS = require('@t-isoyama/exceljs-enhanced');

// Create a workbook
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('My Sheet');

// Add data
worksheet.columns = [
  { header: 'Id', key: 'id', width: 10 },
  { header: 'Name', key: 'name', width: 32 },
];
worksheet.addRow({ id: 1, name: 'John Doe' });
worksheet.addRow({ id: 2, name: 'Jane Doe' });

// Write to file
await workbook.xlsx.writeFile('output.xlsx');

// Read from file
const readWorkbook = new ExcelJS.Workbook();
await readWorkbook.xlsx.readFile('output.xlsx');
const readWorksheet = readWorkbook.getWorksheet('My Sheet');
readWorksheet.eachRow((row, rowNumber) => {
  console.log(`Row ${rowNumber}: ${row.values}`);
});
```

## 📚 Documentation

For comprehensive API documentation, usage examples, and detailed guides, please refer to:

- [Original ExcelJS README](https://github.com/exceljs/exceljs#readme)
- [Original ExcelJS Documentation](https://github.com/exceljs/exceljs)

All APIs from the original ExcelJS library are fully supported and compatible.

## 🔧 Build Commands

```bash
npm run build              # Build the project
npm run clean-build        # Clean and build from scratch
npm test                   # Run full test suite
npm run test:coverage      # Run tests with coverage report
npm run lint               # Run ESLint
npm run lint:fix           # Auto-fix linting issues
```

## 📝 License

This project maintains the same MIT license as the original ExcelJS project.

## 🙏 Credits

This fork is based on the excellent work of:
- Original ExcelJS project and its maintainers
- All contributors to the ExcelJS project

Special thanks to the ExcelJS community for creating and maintaining such a comprehensive Excel library.

## 📊 Performance Improvements

Key performance optimizations included in this fork:

- **Cell Operations**: Pre-computed address cache, optimized coordinate conversions
- **XML Processing**: Lookup tables for encoding, early-exit patterns
- **ZIP Processing**: Migrated from archiver/jszip/unzipper to fflate (faster, smaller)
- **Modern APIs**: Native Node.js APIs instead of polyfills
- **Streaming**: Improved memory efficiency with native streams

See the [commit history](https://github.com/t-isoyama/exceljs-enhanced/commits/master) for detailed changelog.
