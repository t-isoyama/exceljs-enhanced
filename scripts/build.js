/* eslint-disable no-console, import/no-extraneous-dependencies, prefer-template, max-len, consistent-return, no-process-exit */

// Build script for ExcelJS using Browserify + esbuild
// Replaces Grunt + Babel + Terser pipeline
// Browserify for bundling (handles Node.js API polyfills)
// esbuild for minification (100x faster than Terser)

const browserify = require('browserify');
const esbuild = require('esbuild');
const fs = require('fs');
const path = require('path');

console.log('🏗️  Building ExcelJS...\n');

const distDir = path.join(__dirname, '..', 'dist');

// Ensure dist directory exists
if (!fs.existsSync(distDir)) {
  fs.mkdirSync(distDir, {recursive: true});
}

// Browserify helper
function browserifyBundle(entry, output, sourcemap = true) {
  return new Promise((resolve, reject) => {
    const b = browserify(entry, {
      debug: sourcemap, // Enable source maps
      standalone: 'ExcelJS', // Global name for browser
      basedir: path.join(__dirname, '..'), // Set base directory for module resolution
    });

    // Exclude Node.js-only libraries (not available in browser)
    b.exclude('unzipper');
    b.exclude('archiver');

    const writeStream = fs.createWriteStream(output);
    writeStream.on('error', reject);
    writeStream.on('finish', resolve);

    b.bundle((err, buf) => {
      if (err) return reject(err);
      writeStream.write(buf);
      writeStream.end();
    });
  });
}

// Extract inline sourcemap to separate file
function extractSourceMap(jsFile) {
  const content = fs.readFileSync(jsFile, 'utf8');
  const match = content.match(/\/\/# sourceMappingURL=data:application\/json;charset=utf-8;base64,([A-Za-z0-9+/=]+)\s*$/);

  if (match) {
    const base64 = match[1];
    const mapContent = Buffer.from(base64, 'base64').toString('utf8');
    const mapFile = jsFile + '.map';
    fs.writeFileSync(mapFile, mapContent);

    // Replace inline sourcemap with external reference
    const newContent = content.replace(
      match[0],
      `//# sourceMappingURL=${path.basename(mapFile)}`
    );
    fs.writeFileSync(jsFile, newContent);
    return mapFile;
  }
  return null;
}

// Minify using esbuild (much faster than Terser)
async function minify(input, output) {
  const code = fs.readFileSync(input, 'utf8');

  try {
    const result = await esbuild.transform(code, {
      minify: true,
      target: ['chrome90', 'firefox88', 'safari14', 'edge90'],
      sourcemap: true,
      banner: `/*! ExcelJS ${new Date().toISOString().split('T')[0]} | Modern browsers (Chrome 90+, Firefox 88+, Safari 14+) */`,
    });

    fs.writeFileSync(output, result.code);
    if (result.map) {
      fs.writeFileSync(output + '.map', result.map);
    }
  } catch (error) {
    console.error(`Minification error for ${input}:`, error);
    throw error;
  }
}

async function build() {
  try {
    // Build 1: exceljs.js (non-minified)
    console.log('Building exceljs.js (Browserify)...');
    await browserifyBundle('lib/exceljs.browser.js', path.join(distDir, 'exceljs.js'));
    extractSourceMap(path.join(distDir, 'exceljs.js'));
    console.log('✅ exceljs.js');

    // Build 2: exceljs.bare.js (non-minified)
    console.log('Building exceljs.bare.js (Browserify)...');
    await browserifyBundle('lib/exceljs.bare.js', path.join(distDir, 'exceljs.bare.js'));
    extractSourceMap(path.join(distDir, 'exceljs.bare.js'));
    console.log('✅ exceljs.bare.js');

    // Build 3: exceljs.min.js (minified with esbuild)
    console.log('Minifying exceljs.min.js (esbuild)...');
    await minify(
      path.join(distDir, 'exceljs.js'),
      path.join(distDir, 'exceljs.min.js')
    );
    console.log('✅ exceljs.min.js');

    // Build 4: exceljs.bare.min.js (minified with esbuild)
    console.log('Minifying exceljs.bare.min.js (esbuild)...');
    await minify(
      path.join(distDir, 'exceljs.bare.js'),
      path.join(distDir, 'exceljs.bare.min.js')
    );
    console.log('✅ exceljs.bare.min.js');

    // Build 5: Copy LICENSE to dist/
    fs.copyFileSync(
      path.join(__dirname, '..', 'LICENSE'),
      path.join(distDir, 'LICENSE')
    );

    console.log('\n✅ Build complete!\n');
    console.log('📦 Generated files:');
    const files = fs.readdirSync(distDir);
    files
      .filter(f => f.endsWith('.js') || f.endsWith('.map') || f === 'LICENSE')
      .sort()
      .forEach(f => {
        const filePath = path.join(distDir, f);
        if (f !== 'LICENSE') {
          const stats = fs.statSync(filePath);
          const sizeKB = (stats.size / 1024).toFixed(0);
          const sizeMB = (stats.size / 1024 / 1024).toFixed(2);
          const display = stats.size > 1024 * 1024 ? `${sizeMB} MB` : `${sizeKB} KB`;
          console.log(`  ${f.padEnd(30)} ${display.padStart(10)}`);
        } else {
          console.log(`  ${f.padEnd(30)} (license)`);
        }
      });

    console.log('\n🎉 Build completed successfully!');
    console.log('   No Babel, no Grunt, no core-js!');
    console.log('   Browserify for bundling + esbuild for minification = fast & reliable!');
  } catch (error) {
    console.error('\n❌ Build failed:', error);
    process.exit(1);
  }
}

build();
