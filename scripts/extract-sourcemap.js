'use strict';

const fs = require('fs');
const path = require('path');

/**
 * Extract inline source map from JS file and save as separate .map file
 * @param {string} file - Path to JS file with inline source map
 */
function extractSourceMap(file) {
  // eslint-disable-next-line no-console
  console.log(`Processing: ${file}`);

  const content = fs.readFileSync(file, 'utf8');
  const regex = /\/\/# sourceMappingURL=data:application\/json;(?:charset=utf-8;)?base64,([^\s]+)/;
  const match = content.match(regex);

  if (!match) {
    // eslint-disable-next-line no-console
    console.log('  No inline source map found, skipping...');
    return;
  }

  try {
    const base64 = match[1];
    const sourceMapContent = Buffer.from(base64, 'base64').toString('utf8');
    const sourceMapFile = `${file}.map`;

    // Write source map to separate file
    fs.writeFileSync(sourceMapFile, sourceMapContent);
    // eslint-disable-next-line no-console
    console.log(`  Extracted source map: ${path.basename(sourceMapFile)}`);

    // Replace inline source map with reference to external file
    const newContent = content.replace(
      regex,
      `//# sourceMappingURL=${path.basename(sourceMapFile)}`
    );
    fs.writeFileSync(file, newContent);
    // eslint-disable-next-line no-console
    console.log(`  Updated source map reference in ${path.basename(file)}`);
  } catch (error) {
    // eslint-disable-next-line no-console
    console.error(`  Error extracting source map: ${error.message}`);
    throw error;
  }
}

// Process files
const files = [
  './dist/exceljs.js',
  './dist/exceljs.bare.js',
];

// eslint-disable-next-line no-console
console.log('Extracting source maps...\n');

files.forEach(file => {
  if (fs.existsSync(file)) {
    extractSourceMap(file);
  } else {
    // eslint-disable-next-line no-console
    console.log(`File not found: ${file}`);
  }
});

// eslint-disable-next-line no-console
console.log('\nSource map extraction complete!');
