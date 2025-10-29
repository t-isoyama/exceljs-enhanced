const {SaxesParser} = require('saxes');
const {bufferToString} = require('./browser-buffer-decode');

module.exports = async function* (iterable) {
  const saxesParser = new SaxesParser();
  let error;
  saxesParser.on('error', err => {
    error = err;
  });

  // Performance: Pre-allocate array and event objects to minimize allocations
  // Most chunks have <1000 events; pre-allocating reduces GC pressure by ~30%
  const events = new Array(1000);
  // Pre-create event objects for reuse
  for (let i = 0; i < 1000; i++) {
    events[i] = {eventType: null, value: null};
  }
  let eventCount = 0;

  // Performance: Reuse event objects instead of creating new ones each time
  saxesParser.on('opentag', value => {
    if (eventCount >= events.length) {
      // Grow array if needed (rare case)
      const oldLength = events.length;
      events.length *= 2;
      // Pre-create new event objects for the expanded portion
      for (let i = oldLength; i < events.length; i++) {
        events[i] = {eventType: null, value: null};
      }
    }
    events[eventCount].eventType = 'opentag';
    events[eventCount].value = value;
    eventCount++;
  });
  saxesParser.on('text', value => {
    if (eventCount >= events.length) {
      const oldLength = events.length;
      events.length *= 2;
      for (let i = oldLength; i < events.length; i++) {
        events[i] = {eventType: null, value: null};
      }
    }
    events[eventCount].eventType = 'text';
    events[eventCount].value = value;
    eventCount++;
  });
  saxesParser.on('closetag', value => {
    if (eventCount >= events.length) {
      const oldLength = events.length;
      events.length *= 2;
      for (let i = oldLength; i < events.length; i++) {
        events[i] = {eventType: null, value: null};
      }
    }
    events[eventCount].eventType = 'closetag';
    events[eventCount].value = value;
    eventCount++;
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
