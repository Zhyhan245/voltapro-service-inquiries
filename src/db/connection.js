const fs = require('node:fs');
const path = require('node:path');
const { DatabaseSync } = require('node:sqlite');

/**
 * Створює підключення до SQLite з увімкненими зовнішніми ключами та WAL-режимом.
 */
function openDatabase(dbPath) {
  if (dbPath !== ':memory:') {
    const dir = path.dirname(dbPath);
    fs.mkdirSync(dir, { recursive: true });
  }
  const db = new DatabaseSync(dbPath);
  db.function('uk_lower', (val) => (val == null ? '' : String(val).toLowerCase()));
  db.exec('PRAGMA foreign_keys = ON;');
  if (dbPath !== ':memory:') {
    db.exec('PRAGMA journal_mode = WAL;');
  }
  db.exec('PRAGMA busy_timeout = 5000;');
  return db;
}

/**
 * Виконує функцію всередині транзакції SQLite (IMMEDIATE для захисту від конкурентних записів).
 */
function runInTransaction(db, fn) {
  db.exec('BEGIN IMMEDIATE;');
  try {
    const result = fn();
    db.exec('COMMIT;');
    return result;
  } catch (err) {
    try {
      db.exec('ROLLBACK;');
    } catch {
      // Ігноруємо повторну помилку відкату
    }
    throw err;
  }
}

module.exports = { openDatabase, runInTransaction };
