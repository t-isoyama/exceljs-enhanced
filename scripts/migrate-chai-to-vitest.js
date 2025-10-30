
/**
 * Automated Chai to Vitest migration script
 * Converts all Chai assertions to Vitest equivalents
 */

const fs = require('fs');
const path = require('path');
const glob = require('glob');

// Conversion rules - ORDER MATTERS (more specific patterns first)
const transforms = [
  // Chai special constructs that need special handling
  {from: /expect\(([^)]+)\)\.xml\.to\.equal\(/g, to: 'expect($1).toEqualXml('},
  {from: /expect\(([^)]+)\)\.to\.equalDate\(/g, to: 'expect($1).toEqualDate('},

  // .to.be.* patterns (most specific first)
  {from: /\.to\.be\.false\(\)/g, to: '.toBe(false)'},
  {from: /\.to\.be\.true\(\)/g, to: '.toBe(true)'},
  {from: /\.to\.be\.null\(\)/g, to: '.toBeNull()'},
  {from: /\.to\.be\.undefined\(\)/g, to: '.toBeUndefined()'},
  {from: /\.to\.be\.ok\(\)/g, to: '.toBeTruthy()'},
  {from: /\.to\.be\.a\(/g, to: '.toBeTypeOf('},
  {from: /\.to\.be\.an\(/g, to: '.toBeTypeOf('},
  {from: /\.to\.be\.greaterThan\(/g, to: '.toBeGreaterThan('},
  {from: /\.to\.be\.lessThan\(/g, to: '.toBeLessThan('},
  {from: /\.to\.be\.below\(/g, to: '.toBeLessThan('},
  {from: /\.to\.be\.equal\(/g, to: '.toBe('},

  // Negations (must come before non-negated versions)
  {from: /\.to\.not\.equal\(/g, to: '.not.toBe('},
  {from: /\.to\.not\.be\.ok\(\)/g, to: '.not.toBeTruthy()'},
  {from: /\.to\.not\.be\.null\(\)/g, to: '.not.toBeNull()'},
  {from: /\.to\.not\.be\.undefined\(\)/g, to: '.not.toBeUndefined()'},
  {from: /\.to\.not\.throw\(/g, to: '.not.toThrow('},
  {from: /\.to\.not\.throw\(\)/g, to: '.not.toThrow()'},

  // Basic assertions
  {from: /\.to\.deep\.equal\(/g, to: '.toEqual('},
  {from: /\.to\.eql\(/g, to: '.toEqual('},
  {from: /\.to\.equal\(/g, to: '.toBe('},
  {from: /\.to\.throw\(/g, to: '.toThrow('},
  {from: /\.to\.throw\(\)/g, to: '.toThrow()'},

  // Properties and methods
  {from: /\.to\.have\.property\(/g, to: '.toHaveProperty('},
  {from: /\.to\.have\.length\(/g, to: '.toHaveLength('},

  // Other matchers
  {from: /\.to\.match\(/g, to: '.toMatch('},
  {from: /\.to\.exist\b/g, to: '.toBeDefined()'},

  // Clean up chaining artifacts
  {from: /\.to\.toBe\(/g, to: '.toBe('},
  {from: /\.to\.toEqual\(/g, to: '.toEqual('},
];

// Special patterns that need manual review
const manualReviewPatterns = [
  /expect\.fail/,
  /this\.timeout/,
  /this\.slow/,
  /this\.retries/,
];

function transformFile(filePath) {
  let content = fs.readFileSync(filePath, 'utf8');
  let modified = false;
  const warnings = [];

  // Apply transformations
  for (const {from, to} of transforms) {
    const before = content;
    content = content.replace(from, to);
    if (content !== before) {
      modified = true;
    }
  }

  // Check for patterns needing manual review
  for (const pattern of manualReviewPatterns) {
    if (pattern.test(content)) {
      warnings.push(`Pattern ${pattern} found - may need manual review`);
    }
  }

  return {content, modified, warnings};
}

function main() {
  console.log('🔄 Starting Chai to Vitest migration...\n');

  // Find all test files
  const testFiles = glob.sync('spec/**/*.spec.js', {
    cwd: path.join(__dirname, '..'),
    absolute: true,
  });

  console.log(`Found ${testFiles.length} test files\n`);

  let modifiedCount = 0;
  const filesWithWarnings = [];

  for (const filePath of testFiles) {
    const {content, modified, warnings} = transformFile(filePath);

    if (modified) {
      fs.writeFileSync(filePath, content, 'utf8');
      modifiedCount++;
      console.log(`✅ ${path.relative(process.cwd(), filePath)}`);
    }

    if (warnings.length > 0) {
      filesWithWarnings.push({
        file: path.relative(process.cwd(), filePath),
        warnings,
      });
    }
  }

  console.log('\n📊 Summary:');
  console.log(`  Modified: ${modifiedCount}/${testFiles.length} files`);

  if (filesWithWarnings.length > 0) {
    console.log(`\n⚠️  Files needing manual review (${filesWithWarnings.length}):`);
    for (const {file, warnings} of filesWithWarnings) {
      console.log(`\n  ${file}:`);
      for (const warning of warnings) {
        console.log(`    - ${warning}`);
      }
    }
  }

  console.log('\n✨ Migration complete!');
}

try {
  main();
} catch (error) {
  console.error('❌ Migration failed:', error);
  process.exit(1);
}
