# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [4.4.0-enhanced.0] - 2025-11-05

### Overview
This is a performance-optimized fork of [ExcelJS](https://github.com/exceljs/exceljs) v4.4.0, maintaining full API compatibility while delivering significant performance improvements and modernization.

### Performance Improvements
- **7-18x cumulative speedup** through algorithmic optimizations
- Pre-computed lookup tables for XML encoding and column cache
- Early-exit patterns throughout the codebase
- Optimized data structures and algorithms

### Modernization
- **Updated Dependencies**:
  - Replaced `archiver`/`jszip`/`unzipper` with `fflate` for unified ZIP operations
  - Replaced `sax` with `saxes` for modern SAX parsing
  - Replaced `moment` with `dayjs` for date handling
  - Replaced Mocha with Vitest for testing
- **Build System**: Replaced Grunt + Babel with Browserify + esbuild
- **Node.js**: Minimum version upgraded to 22+ with native API support (Object.hasOwn, native streams)
- **Browser Bundle**: Optimized bundle size with modern build tools

### Security
- All dependencies updated to latest versions
- Resolved known security vulnerabilities from upstream

### Breaking Changes
- **Node.js 22+ Required**: This version requires Node.js 22 or higher
- No other breaking changes to the API - fully compatible with ExcelJS v4.4.0

### Migration from Original ExcelJS
Simply replace `exceljs` with `@t-isoyama/exceljs-enhanced` in your package.json:

```bash
npm uninstall exceljs
npm install @t-isoyama/exceljs-enhanced
```

No code changes required - the API is 100% compatible.

### Acknowledgments
This project is based on the excellent work by [Guyon Roche](https://github.com/guyonroche) and the ExcelJS contributors. All credit for the original design and implementation goes to them.

---

## Version History

For the original ExcelJS changelog, see: https://github.com/exceljs/exceljs/blob/master/CHANGELOG.md

[4.4.0-enhanced.0]: https://github.com/t-isoyama/exceljs-enhanced/releases/tag/v4.4.0-enhanced.0
