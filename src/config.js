const path = require('node:path');

/**
 * Централізована конфігурація застосунку з параметрами середовища.
 */
function loadConfig(overrides = {}) {
  const port = Number(overrides.port ?? process.env.PORT ?? 3000);
  const rawDbPath = overrides.dbPath ?? process.env.DB_PATH ?? path.join(process.cwd(), 'data', 'voltapro.db');
  const dbPath = rawDbPath === ':memory:' ? ':memory:' : path.resolve(rawDbPath);
  const sessionTtlHours = Number(overrides.sessionTtlHours ?? process.env.SESSION_TTL_HOURS ?? 12);
  const nodeEnv = overrides.nodeEnv ?? process.env.NODE_ENV ?? 'development';

  return {
    port: Number.isFinite(port) && port > 0 ? port : 3000,
    dbPath,
    sessionTtlHours: Number.isFinite(sessionTtlHours) && sessionTtlHours > 0 ? sessionTtlHours : 12,
    nodeEnv,
    cookieName: 'vp_session'
  };
}

module.exports = { loadConfig };
