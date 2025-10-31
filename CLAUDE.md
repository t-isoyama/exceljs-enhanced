# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**ExcelJS Enhanced** is a performance-optimized fork of [ExcelJS](https://github.com/exceljs/exceljs) v4.4.0. It provides comprehensive Excel Workbook management for Node.js and browsers with significant performance improvements and modernization while maintaining full API compatibility.

### This is a Fork

- **Upstream**: [exceljs/exceljs](https://github.com/exceljs/exceljs)
- **Fork Repository**: [t-isoyama/exceljs-enhanced](https://github.com/t-isoyama/exceljs-enhanced)
- **Base Version**: v4.4.0

### Key Changes from Upstream

1. **Performance Optimizations**: 7-18x cumulative speedup through optimized algorithms
2. **Modern Dependencies**: fflate (ZIP), saxes (XML), dayjs (dates), vitest (testing)
3. **Build System**: Browserify + esbuild (replaced Grunt + Babel)
4. **Node.js Version**: Minimum Node.js 22+ (uses modern native APIs)
5. **Security**: Updated all dependencies to fix vulnerabilities

## Build and Test Commands

```bash
# Building
npm run build                    # Build using Browserify + esbuild
npm run clean-build              # Clean and build from scratch
npm run clean                    # Remove build artifacts

# Testing
npm test                         # Run full test suite (unit + integration + end-to-end)
npm run test:unit                # Run unit tests only
npm run test:integration         # Run integration tests only
npm run test:end-to-end          # Run end-to-end tests only
npm run test:coverage            # Run tests with coverage report
npm run test:watch               # Run tests in watch mode
npm run test:ui                  # Open Vitest UI for interactive testing

# Linting
npm run lint                     # Run ESLint
npm run lint:fix                 # Auto-fix linting issues

# Single Test File
vitest run spec/unit/path/to/test.spec.js
vitest watch spec/unit/path/to/test.spec.js  # Watch mode
```

## Architecture Overview

### Core Module Structure

#### 1. Document Model Layer (`lib/doc/`)
In-memory representation of Excel files:
- `workbook.js` - Root container managing worksheets, defined names, media
- `worksheet.js` - Manages cells, rows, columns, merged cells, validations
- `row.js`, `column.js`, `cell.js` - Low-level data structures
- `range.js` - Cell range operations

#### 2. Streaming Layer (`lib/stream/xlsx/`)
Memory-efficient operations for large files:
- `workbook-writer.js`, `workbook-reader.js` - Streaming read/write
- `worksheet-writer.js`, `worksheet-reader.js` - Worksheet-level streaming

#### 3. XLSX Transform Layer (`lib/xlsx/xform/`)
XML serialization/deserialization using SAX parsing:
- `book/` - Workbook XML transforms
- `sheet/` - Worksheet XML transforms
- `style/` - Style XML transforms
- `core/` - Office document core transforms

Each transform extends `BaseXform` with `render()` and `parseOpen()`/`parseText()`/`parseClose()` methods.

#### 4. CSV Layer (`lib/csv/`)
CSV processing using fast-csv library.

#### 5. Utility Layer (`lib/utils/`)
Common utilities:
- `col-cache.js` - Column letter-number conversion (optimized)
- `xml-stream.js` - XML generation
- `zip-stream.js` - ZIP handling using fflate
- `parse-sax.js` - SAX-based XML parsing
- `under-dash.js` - Minimal lodash-like utilities

### Key Design Patterns

1. **Document vs. Streaming Mode**
   - Document: Full in-memory model (small/medium files)
   - Streaming: Sequential processing (large files)

2. **Transform Pattern**
   All XLSX operations use consistent transform pattern via `BaseXform`

3. **Performance Optimizations**
   - Pre-computed lookup tables (XML encoding, column cache)
   - Early-exit patterns
   - Native Node.js APIs (Object.hasOwn, native streams)

## Important Implementation Notes

### ZIP Processing
- **Current**: Uses `fflate` for unified Node.js/browser ZIP operations
- **Old**: archiver/jszip/unzipper (removed for performance/size)

### XML Parsing
- **Current**: Uses `saxes` (modern SAX parser)
- **Old**: `sax` (removed)

### Date Handling
- Uses `dayjs` instead of moment
- Supports both 1900 and 1904 date systems

### Testing Framework
- **Current**: Vitest (fast, modern)
- **Old**: Mocha (removed)

## File Structure

```
lib/                 # Source code (ES6+)
  ├── doc/           # Document model
  ├── stream/        # Streaming operations
  ├── xlsx/          # XLSX format handling
  ├── csv/           # CSV format handling
  └── utils/         # Utility functions
dist/                # Built browser bundles
spec/                # Test files
  ├── unit/          # Unit tests
  ├── integration/   # Integration tests
  └── end-to-end/    # End-to-end tests
scripts/             # Build scripts
excel.js             # Main entry point
index.d.ts           # TypeScript definitions
```

## Common Tasks

### Adding a New Feature
1. Implement in `lib/` with proper module structure
2. Add corresponding xform if it affects XLSX format
3. Add unit tests in `spec/unit/`
4. Add integration tests in `spec/integration/`
5. Update TypeScript definitions in `index.d.ts`

### Performance Optimization
- Use pre-computed lookup tables where possible
- Implement early-exit patterns
- Avoid repeated regex compilation
- Use native APIs (Object.hasOwn, etc.)
- Profile with `npm run benchmark`

### Debugging
- Run specific test: `vitest run spec/unit/path/to/test.spec.js`
- Use Vitest UI: `npm run test:ui`
- Check coverage: `npm run test:coverage`

## Code Style

- ESLint with Airbnb base config + Prettier
- Automatic formatting on commit (husky + lint-staged)
- Modern JavaScript (ES6+, async/await)
- No unnecessary polyfills (Node.js 22+ target)

## TypeScript Support

- Type definitions in `index.d.ts`
- Must stay in sync with implementation
- Run `npm run test:typescript` to validate

## Important: This is a Fork

When making changes:
- Maintain API compatibility with original ExcelJS
- Focus on performance and modernization
- Document significant changes
- Keep tests comprehensive

For API documentation, refer to [original ExcelJS documentation](https://github.com/exceljs/exceljs).
