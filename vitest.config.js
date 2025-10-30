import {defineConfig} from 'vitest/config';
import {resolve} from 'path';

export default defineConfig({
  test: {
    // Enable global APIs (describe, it, expect, etc.)
    globals: true,

    // Test environment
    environment: 'node',

    // Setup files to run before each test file
    setupFiles: ['./spec/config/vitest-setup.js'],

    // Test file patterns
    include: [
      'spec/unit/**/*.spec.js',
      'spec/integration/**/*.spec.js',
      'spec/end-to-end/**/*.spec.js',
      'spec/typescript/**/*.spec.ts',
    ],

    // Exclude patterns
    exclude: [
      'node_modules',
      'dist',
      'build',
      'spec/browser/**',
      'spec/manual/**',
    ],

    // Test timeout (2 minutes default, matching Mocha)
    testTimeout: 120000,

    // Hook timeout
    hookTimeout: 10000,

    // Parallel execution configuration
    maxConcurrency: 5,
    minWorkers: 1,
    maxWorkers: undefined, // Use all available CPUs

    // Reporter configuration
    reporters: ['verbose'],

    // Coverage configuration
    coverage: {
      provider: 'v8',
      reporter: ['text', 'json', 'html', 'lcov'],
      include: ['lib/**/*.js'],
      exclude: [
        'node_modules',
        'spec',
        'dist',
        'build',
      ],
      all: true,
      lines: 80,
      functions: 80,
      branches: 80,
      statements: 80,
    },

    // Retry failed tests
    retry: 0,

    // Pool options
    pool: 'forks',
    poolOptions: {
      forks: {
        singleFork: false,
      },
    },
  },

  // Resolve configuration for verquire compatibility
  resolve: {
    alias: {
      // Allow verquire to resolve lib/ modules
      lib: resolve(__dirname, './lib'),
    },
  },

  // ESM/CommonJS interop
  esbuild: {
    target: 'node22',
  },
});
