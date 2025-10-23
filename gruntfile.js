'use strict';

module.exports = function(grunt) {
  grunt.loadNpmTasks('grunt-babel');
  grunt.loadNpmTasks('grunt-browserify');
  grunt.loadNpmTasks('grunt-terser');
  grunt.loadNpmTasks('grunt-contrib-jasmine');
  grunt.loadNpmTasks('grunt-contrib-copy');

  grunt.initConfig({
    babel: {
      options: {
        sourceMap: true,
        compact: false,
      },
      dist: {
        files: [
          {
            expand: true,
            src: ['./lib/**/*.js', './spec/browser/*.js'],
            dest: './build/',
          },
        ],
      },
    },
    browserify: {
      options: {
        transform: [
          [
            'babelify',
            {
              // enable babel transpile for node_modules
              global: true,
              presets: ['@babel/preset-env'],
              // core-js should not be transpiled
              // See https://github.com/zloirock/core-js/issues/514
              ignore: [/node_modules[\\/]core-js/],
            },
          ],
        ],
        browserifyOptions: {
          // enable source map for browserify
          debug: true,
          standalone: 'ExcelJS',
        },
      },
      bare: {
        // keep the original source for source maps
        src: ['./lib/exceljs.bare.js'],
        dest: './dist/exceljs.bare.js',
      },
      bundle: {
        // keep the original source for source maps
        src: ['./lib/exceljs.browser.js'],
        dest: './dist/exceljs.js',
      },
      spec: {
        options: {
          transform: null,
          browserifyOptions: null,
        },
        src: ['./build/spec/browser/exceljs.spec.js'],
        dest: './build/web/exceljs.spec.js',
      },
    },

    terser: {
      options: {
        output: {
          preamble: '/*! ExcelJS <%= grunt.template.today("dd-mm-yyyy") %> */\n',
          ascii_only: true,
        },
      },
      dist: {
        options: {
          // Keep the original source maps from browserify
          // See also https://www.npmjs.com/package/terser#source-map-options
          sourceMap: {
            content: 'inline',
            url: 'exceljs.min.js.map',
          },
        },
        files: {
          './dist/exceljs.min.js': ['./dist/exceljs.js'],
        },
      },
      bare: {
        options: {
          // Keep the original source maps from browserify
          // See also https://www.npmjs.com/package/terser#source-map-options
          sourceMap: {
            content: 'inline',
            url: 'exceljs.bare.min.js.map',
          },
        },
        files: {
          './dist/exceljs.bare.min.js': ['./dist/exceljs.bare.js'],
        },
      },
    },

    copy: {
      dist: {
        files: [
          {expand: true, src: ['**'], cwd: './build/lib', dest: './dist/es5'},
          {src: './build/lib/exceljs.nodejs.js', dest: './dist/es5/index.js'},
          {src: './LICENSE', dest: './dist/LICENSE'},
        ],
      },
    },

    jasmine: {
      options: {
        noSandbox: true,
        timeout: 30000,
      },
      dev: {
        src: ['./dist/exceljs.js'],
        options: {
          specs: './build/web/exceljs.spec.js',
        },
      },
    },
  });

  // Custom task to extract source maps
  grunt.registerTask('extract-sourcemap', 'Extract inline source maps to separate files', function() {
    const done = this.async();
    const {execFile} = require('child_process');

    execFile('node', ['./scripts/extract-sourcemap.js'], (error, stdout, stderr) => {
      if (stdout) {
        grunt.log.writeln(stdout);
      }
      if (stderr) {
        grunt.log.error(stderr);
      }
      if (error) {
        grunt.fail.fatal(`Source map extraction failed: ${error.message}`);
        done(false);
      } else {
        done(true);
      }
    });
  });

  grunt.registerTask('build', ['babel:dist', 'browserify', 'terser', 'extract-sourcemap', 'copy']);
  grunt.registerTask('ug', ['terser']);
};
