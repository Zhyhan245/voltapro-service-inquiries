const crypto = require('node:crypto');

/**
 * Структурований логер без виведення паролів, токенів або приватних форм.
 */
function createRequestId() {
  return crypto.randomBytes(6).toString('hex');
}

function logEvent(level, component, message, meta = {}) {
  const entry = {
    ts: new Date().toISOString(),
    level,
    component,
    message,
    ...meta
  };
  if (level === 'error') {
    console.error(JSON.stringify(entry));
  } else {
    console.log(JSON.stringify(entry));
  }
}

module.exports = { createRequestId, logEvent };
