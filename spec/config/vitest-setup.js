// Vitest global setup file
// Complete migration from Mocha/Chai to pure Vitest

import { expect, beforeAll, afterAll, beforeEach, afterEach, describe } from 'vitest';

// Import verquire for global access
const verquire = require('../utils/verquire');

// Make verquire globally available (used in many test files)
global.verquire = verquire;

// Add Mocha-style global hooks for backward compatibility
global.before = beforeAll;
global.after = afterAll;
global.beforeEach = beforeEach;
global.afterEach = afterEach;

// Add Mocha-style context alias for describe
global.context = describe;

// Make Vitest's expect available globally
global.expect = expect;

// ============================================================================
// Custom Vitest Matchers - Replacing Chai plugins
// ============================================================================

expect.extend({
  // ========================================
  // XML Matcher (replaces chai-xml)
  // ========================================
  toEqualXml(received, expected) {
    // Sort attributes within XML tags to make comparison order-independent
    const sortAttributes = xml => {
      return String(xml).replace(/<(\w+)([^>\/]*)(\/?)/g, (match, tag, attrs, closing) => {
        if (!attrs.trim()) {
          return `<${tag}${closing}`;
        }
        // Extract and sort attribute key-value pairs
        const attrPairs = attrs.trim().match(/(\w+)="([^"]*)"/g) || [];
        const sorted = attrPairs.sort().join(' ');
        return `<${tag} ${sorted}${closing}`;
      });
    };

    // Normalize XML strings by:
    // 1. Sorting attributes within each tag
    // 2. Removing whitespace between tags (> <)
    // 3. Removing leading/trailing whitespace per line
    // 4. Normalizing internal whitespace to single space
    // 5. Normalizing self-closing tags (remove space before /> and >)
    // 6. Trimming
    const normalize = xml =>
      sortAttributes(String(xml))
        .replace(/>\s+</g, '><')
        .replace(/^\s+|\s+$/gm, '')
        .replace(/\s+/g, ' ')
        .replace(/\s+\/>/g, '/>') // Normalize self-closing tags
        .replace(/\s+>/g, '>') // Remove space before >
        .replace(/<\/>/g, '/>')  // Fix malformed closing tags
        .trim();

    const normalizedReceived = normalize(received);
    const normalizedExpected = normalize(expected);

    const pass = normalizedReceived === normalizedExpected;

    return {
      pass,
      message: () =>
        pass
          ? `Expected XML not to be equal`
          : `Expected XML to be equal\n\nReceived:\n${normalizedReceived}\n\nExpected:\n${normalizedExpected}`,
    };
  },

  // ========================================
  // Date Matcher (replaces chai-datetime)
  // ========================================
  toEqualDate(received, expected, tolerance = 0) {
    const receivedTime = received instanceof Date ? received.getTime() : NaN;
    const expectedTime = expected instanceof Date ? expected.getTime() : NaN;

    if (Number.isNaN(receivedTime) || Number.isNaN(expectedTime)) {
      return {
        pass: false,
        message: () => `Expected both values to be valid Date objects\nReceived: ${received}\nExpected: ${expected}`,
      };
    }

    const diff = Math.abs(receivedTime - expectedTime);
    const pass = diff <= tolerance;

    return {
      pass,
      message: () =>
        pass
          ? `Expected dates not to be equal (tolerance: ${tolerance}ms)`
          : `Expected dates to be equal (tolerance: ${tolerance}ms)\nReceived: ${received.toISOString()}\nExpected: ${expected.toISOString()}\nDifference: ${diff}ms`,
    };
  },

  // ========================================
  // Additional Date Matchers
  // ========================================
  toBeDate(received) {
    const pass = received instanceof Date && !Number.isNaN(received.getTime());
    return {
      pass,
      message: () =>
        pass
          ? `Expected value not to be a valid Date`
          : `Expected value to be a valid Date, received: ${typeof received}`,
    };
  },

  toBeBeforeDate(received, expected) {
    const receivedTime = received instanceof Date ? received.getTime() : NaN;
    const expectedTime = expected instanceof Date ? expected.getTime() : NaN;
    const pass = !Number.isNaN(receivedTime) && !Number.isNaN(expectedTime) && receivedTime < expectedTime;

    return {
      pass,
      message: () =>
        pass
          ? `Expected ${received.toISOString()} not to be before ${expected.toISOString()}`
          : `Expected ${received.toISOString()} to be before ${expected.toISOString()}`,
    };
  },

  toBeAfterDate(received, expected) {
    const receivedTime = received instanceof Date ? received.getTime() : NaN;
    const expectedTime = expected instanceof Date ? expected.getTime() : NaN;
    const pass = !Number.isNaN(receivedTime) && !Number.isNaN(expectedTime) && receivedTime > expectedTime;

    return {
      pass,
      message: () =>
        pass
          ? `Expected ${received.toISOString()} not to be after ${expected.toISOString()}`
          : `Expected ${received.toISOString()} to be after ${expected.toISOString()}`,
    };
  },
});
