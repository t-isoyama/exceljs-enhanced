# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

ExcelJS is a comprehensive Excel Workbook Manager library for Node.js and browsers that enables reading, manipulating, and writing spreadsheet data in XLSX and CSV formats. The library is reverse-engineered from Excel spreadsheet files and supports a wide range of Excel features including styles, formulas, data validations, conditional formatting, images, and pivot tables.

## Build and Test Commands

### Building
```bash
npm run build                    # Build the project using Grunt
npm run clean-build              # Clean and build from scratch
npm run clean                    # Remove build artifacts
```

### Testing
```bash
npm test                         # Run full test suite (unit + integration + end-to-end)
npm run test:unit                # Run unit tests only
npm run test:integration         # Run integration tests only
npm run test:end-to-end          # Run end-to-end tests only
npm run test:typescript          # Run TypeScript tests
npm run test:dist                # Test distribution files
npm run test:watch               # Run tests in watch mode
npm run test:ui                  # Open Vitest UI for interactive testing
npm run test:coverage            # Run tests with coverage report
```

### Linting
```bash
npm run lint                     # Run ESLint
npm run lint:fix                 # Auto-fix linting issues with prettier-eslint
```

### Running Tests for a Single Test File
```bash
vitest run spec/unit/path/to/test.spec.js
vitest run spec/integration/path/to/test.spec.js
vitest watch spec/unit/path/to/test.spec.js  # Watch mode for specific file
```

## Architecture Overview

### Core Module Structure

ExcelJS uses a modular architecture divided into several key subsystems:

#### 1. Document Model Layer (`lib/doc/`)
The document model represents the in-memory structure of Excel files:

- **`workbook.js`**: Root container managing worksheets, defined names, media, pivot tables
- **`worksheet.js`**: Manages cells, rows, columns, merged cells, data validations, conditional formatting, tables
- **`row.js`**: Row-level operations including cell access, styling, outlines
- **`column.js`**: Column-level operations including width, styling, outlines
- **`cell.js`**: Individual cell management with values, formulas, styles, hyperlinks, notes
- **`range.js`**: Cell range operations and addressing
- **`table.js`**: Excel table functionality with filtering and totals
- **`pivot-table.js`**: Pivot table structure and configuration
- **`data-validations.js`**: Cell-level data validation rules
- **`defined-names.js`**: Workbook-level named ranges

#### 2. Streaming Layer (`lib/stream/xlsx/`)
Provides memory-efficient streaming read/write operations:

- **`workbook-writer.js`**: Streaming writer for generating large XLSX files
- **`workbook-reader.js`**: Streaming reader using async iterators for parsing XLSX
- **`worksheet-writer.js`**: Worksheet-level streaming write operations
- **`worksheet-reader.js`**: Worksheet-level streaming read operations
- **`hyperlink-reader.js`**: Specialized hyperlink parsing
- **`sheet-rels-writer.js`**: Relationship XML generation
- **`sheet-comments-writer.js`**: Comment XML generation

#### 3. XLSX Transform Layer (`lib/xlsx/xform/`)
Handles XML serialization/deserialization for XLSX format:

- **`book/`**: Workbook-level transforms (workbook.xml, properties)
- **`sheet/`**: Worksheet-level transforms (sheet XML, cells, rows, columns)
- **`style/`**: Style transforms (fonts, fills, borders, number formats)
- **`drawing/`**: Image and drawing transforms
- **`table/`**: Table transforms
- **`pivot-table/`**: Pivot table and cache transforms
- **`comment/`**: Cell comment transforms (XML and VML)
- **`core/`**: Core Office document transforms (relationships, content-types)

Each transform extends `BaseXform` and implements `render()` and `parseOpen()`/`parseText()`/`parseClose()` methods for XML processing.

#### 4. CSV Layer (`lib/csv/`)
CSV file reading and writing using the fast-csv library:

- **`csv.js`**: Main CSV interface
- **`line-buffer.js`**: Buffer management for line-by-line processing
- **`stream-converter.js`**: Stream conversion utilities

#### 5. Utility Layer (`lib/utils/`)
Common utilities used throughout the codebase:

- **`col-cache.js`**: Column letter-to-number conversion caching
- **`xml-stream.js`**: XML generation utilities
- **`zip-stream.js`**: ZIP archive handling
- **`stream-buf.js`**: In-memory stream buffer
- **`shared-strings.js`**: Shared string table management
- **`shared-formula.js`**: Shared formula translation
- **`cell-matrix.js`**: 2D cell grid management
- **`encryptor.js`**: Worksheet protection encryption
- **`copy-style.js`**: Style object deep copying
- **`parse-sax.js`**: SAX-based XML parsing
- **`under-dash.js`**: Lodash-like utility functions

### Key Design Patterns

#### 1. Document vs. Streaming Mode
ExcelJS supports two modes of operation:

- **Document Mode** (`lib/doc/workbook.js`): Full in-memory model allowing random access and modification. Suitable for small to medium files.
- **Streaming Mode** (`lib/stream/xlsx/workbook-writer.js`, `workbook-reader.js`): Memory-efficient sequential processing using Node.js streams and async iterators. Required for very large files.

#### 2. Transform Pattern
All XLSX XML serialization uses a consistent transform pattern via `BaseXform`:
- `prepare()`: Pre-processing before serialization
- `render()`: Generate XML for writing
- `parseOpen()`: Handle XML element opening
- `parseText()`: Handle text content
- `parseClose()`: Handle XML element closing
- `reconcile()`: Post-processing after parsing

#### 3. Cell Addressing
Cells use two addressing systems:
- **A1 notation**: e.g., "A1", "B5", "AA100"
- **Row/Column indices**: 1-based for user API, 0-based internally in arrays

The `colCache` utility provides efficient conversion between column letters and numbers.

#### 4. Style Management
Styles are shared objects:
- When a style is assigned to a cell/row/column, they share the same object reference
- Modifying a style object affects all entities referencing it
- Use `copyStyle()` utility for independent style copies
- Style properties: `numFmt`, `font`, `alignment`, `border`, `fill`, `protection`

### Data Flow

#### Reading XLSX Files (Document Mode)
1. `workbook.xlsx.readFile()` or `workbook.xlsx.read()`
2. Unzip XLSX container using JSZip or unzipper
3. Parse XML files using SAX parser (saxes)
4. Transform XML to model objects via xform classes
5. Build in-memory workbook/worksheet/cell structure
6. Return completed Workbook instance

#### Writing XLSX Files (Document Mode)
1. `workbook.xlsx.writeFile()` or `workbook.xlsx.write()`
2. Transform model objects to XML via xform classes
3. Generate XML streams for each component (workbook.xml, sheets, styles, etc.)
4. Package XML files into ZIP archive using archiver
5. Write or return stream/buffer

#### Streaming Read (Async Iterator Pattern)
```javascript
const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filename);
for await (const worksheetReader of workbookReader) {
  for await (const row of worksheetReader) {
    // Process row
  }
}
```

#### Streaming Write (Commit Pattern)
```javascript
const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
const sheet = workbook.addWorksheet('Sheet1');
const row = sheet.addRow([1, 2, 3]);
row.commit(); // Commit row to stream
sheet.commit(); // Commit worksheet
await workbook.commit(); // Finalize workbook
```

## Important Implementation Details

### Cell Value Types
Cells support multiple value types (see `lib/doc/cell.js`):
- Null, Number, String, Date, Boolean, Error
- Hyperlink (text + URL)
- Formula (with optional result and shared formula support)
- Rich Text (in-cell formatting)

### Formulas
- **Shared formulas**: Master cell contains formula, slave cells reference it with translation
- **Array formulas**: Single formula applied to a range without translation
- Formula results must be provided; ExcelJS does not evaluate formulas

### Merged Cells
- Merge operations link cells to a master cell
- Merged cells share style with master
- Splicing rows/columns with merged cells can produce unpredictable results

### Tables
- Tables are Excel's structured data format with headers, data, and optional totals row
- Support for filtering, sorting, totals functions, and styling
- Adding a table modifies the worksheet by inserting headers and data

### Images
- Images are added to workbook first (returns imageId)
- Image can be added as tiled background or positioned over cell range
- Supports anchoring modes: absolute, oneCell (move with cells), twoCells (move and size)
- Image formats: JPEG, PNG, GIF

### Conditional Formatting
- Supports expression-based, cell comparison, top/bottom, above/below average
- Icon sets, color scales, and data bars
- Priority determines precedence when rules overlap
- ExtLst-based rules (dataBar, some icon sets) fully supported

### Data Validations
- Types: list, whole, decimal, textLength, date, time, custom
- Operators: between, notBetween, equal, notEqual, greaterThan, lessThan, etc.
- Can display input messages and error alerts

### Protection
- Worksheet protection with password (uses encryption with configurable spinCount)
- Cell-level protection (locked/hidden) only effective when sheet is protected
- Protection does not encrypt file content, only prevents editing in Excel

### Themes
- Theme files from original XLSX are preserved through read/write cycle
- Use `workbook.clearThemes()` to remove theme files
- Default theme (theme1.xml) is included in generated files

## File Structure Notes

- **`excel.js`**: Main entry point exporting the Workbook class
- **`index.d.ts`**: TypeScript type definitions
- **`dist/`**: Browserified and transpiled distributions
  - `exceljs.js`: Bundle with polyfills
  - `exceljs.bare.js`: Bundle without polyfills
  - `exceljs.min.js`: Minified browser bundle
  - `es5/`: ES5 transpiled code for older Node.js versions
- **`lib/`**: Original source code (ES6+)
- **`spec/`**: Test files organized as:
  - `unit/`: Unit tests for individual modules
  - `integration/`: Integration tests for workflows
  - `end-to-end/`: End-to-end tests with file I/O
  - `browser/`: Browser environment tests

## Common Pitfalls and Solutions

### Issue: Out of Memory with Large Files
**Solution**: Use streaming API (`stream.xlsx.WorkbookWriter` / `WorkbookReader`) instead of document API.

### Issue: Incorrect Cell References After Splicing
**Solution**: Splicing affects defined names and merged cells. Update references manually or avoid splicing when possible.

### Issue: Shared Formula Translation Issues
**Solution**: Ensure master formula cell exists before slave cells. Use absolute references ($A$1) where translation is not desired.

### Issue: Style Changes Affect Multiple Cells
**Solution**: Styles are shared by reference. Clone style objects using spread operator or `copyStyle()` utility before modification.

### Issue: Date Formatting/Parsing Issues
**Solution**: Check `workbook.properties.date1904` flag. Excel on Mac uses 1904 date system by default.

### Issue: Cell Comments Not Visible
**Solution**: Comments require both comment content and VML drawing. Ensure both are properly generated.

### Issue: Conditional Formatting Not Working
**Solution**: Ensure formula references use correct cell address. Check priority values don't conflict.

## Testing Conventions

- Unit tests use Vitest with Chai compatibility layer
- Test files mirror source structure: `lib/doc/worksheet.js` → `spec/unit/doc/worksheet.spec.js`
- Integration tests in `spec/integration/` often test specific GitHub issues
- Use `spec/config/vitest-setup.js` for test environment configuration
- Global test APIs (describe, it, expect) are automatically available
- Custom matchers for XML comparison and date assertions included
- Browser tests can use Vitest Browser Mode (optional)

## TypeScript Support

- Type definitions maintained in `index.d.ts`
- TypeScript compilation tests ensure types stay in sync with implementation
- Run `npm run test:typescript` to validate type definitions

## Major Version 4 Changes

- Migrated from streams to async iterators for cleaner code
- Removed Promise library dependency injection (uses native Promises)
- Main export is now original ES6+ source (was transpiled ES5)
- ES5 builds still available in `dist/es5/` for compatibility
