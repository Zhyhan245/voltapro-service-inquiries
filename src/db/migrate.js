const fs = require('node:fs');
const path = require('node:path');
const { loadConfig } = require('../config');
const { openDatabase, runInTransaction } = require('./connection');

/**
 * Застосовує всі SQL-міграції з каталогу migrations/ у порядку версій.
 * Підтримує повторний запуск без дублювання (ідемпотентність).
 */
function runMigrations(db, migrationsDir = path.join(__dirname, '..', '..', 'migrations')) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS schema_migrations (
      version TEXT PRIMARY KEY,
      applied_at TEXT NOT NULL
    );
  `);

  const files = fs
    .readdirSync(migrationsDir)
    .filter((file) => file.endsWith('.sql'))
    .sort();

  const selectApplied = db.prepare('SELECT 1 FROM schema_migrations WHERE version = ?');
  const insertApplied = db.prepare(
    'INSERT INTO schema_migrations (version, applied_at) VALUES (?, ?)'
  );

  const appliedNow = [];
  for (const file of files) {
    const exists = selectApplied.get(file);
    if (exists) continue;

    const sql = fs.readFileSync(path.join(migrationsDir, file), 'utf8');
    runInTransaction(db, () => {
      db.exec(sql);
      insertApplied.run(file, new Date().toISOString());
    });
    appliedNow.push(file);
  }

  return appliedNow;
}

if (require.main === module) {
  const config = loadConfig();
  const db = openDatabase(config.dbPath);
  try {
    const applied = runMigrations(db);
    console.log(
      applied.length > 0
        ? `Застосовано міграції: ${applied.join(', ')}`
        : 'Усі міграції вже застосовано.'
    );
  } finally {
    db.close();
  }
}

module.exports = { runMigrations };
