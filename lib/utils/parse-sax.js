const {SaxesParser} = require('saxes');
const {PassThrough} = require('readable-stream');
const {bufferToString} = require('./browser-buffer-decode');

module.exports = async function* (iterable) {
  // TODO: Remove once node v8 is deprecated
  // Detect and upgrade old streams
  if (iterable.pipe && !iterable[Symbol.asyncIterator]) {
    iterable = iterable.pipe(new PassThrough());
  }
  const saxesParser = new SaxesParser();
  let error;
  saxesParser.on('error', err => {
    error = err;
  });

  // Performance: Pre-allocate array with reasonable size to avoid repeated reallocation
  // Most chunks will have fewer than 1000 events, so this is a good balance
  const events = new Array(1000);
  let eventCount = 0;

  // Performance: Use index-based insertion instead of push for better performance
  saxesParser.on('opentag', value => {
    if (eventCount >= events.length) {
      // Grow array if needed (rare case)
      events.length *= 2;
    }
    events[eventCount++] = {eventType: 'opentag', value};
  });
  saxesParser.on('text', value => {
    if (eventCount >= events.length) {
      events.length *= 2;
    }
    events[eventCount++] = {eventType: 'text', value};
  });
  saxesParser.on('closetag', value => {
    if (eventCount >= events.length) {
      events.length *= 2;
    }
    events[eventCount++] = {eventType: 'closetag', value};
  });

  for await (const chunk of iterable) {
    saxesParser.write(bufferToString(chunk));
    // saxesParser.write and saxesParser.on() are synchronous,
    // so we can only reach the below line once all events have been emitted
    if (error) throw error;
    // As a performance optimization, we gather all events instead of passing
    // them one by one, which would cause each event to go through the event queue

    // Performance: Only yield the filled portion of the array
    if (eventCount > 0) {
      yield events.slice(0, eventCount);
      eventCount = 0; // Reset counter for reuse
    }
  }
};
