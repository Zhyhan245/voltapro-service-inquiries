/**
 * Репозиторій виробничого календаря (святкові та перенесені робочі дні).
 */
function createCalendarRepository(db) {
  function getExceptionsMap() {
    const rows = db
      .prepare('SELECT date, is_working_day FROM calendar_exceptions')
      .all();
    const map = new Map();
    for (const row of rows) {
      map.set(row.date, row.is_working_day === 1);
    }
    return map;
  }

  function listExceptions() {
    return db
      .prepare('SELECT date, is_working_day, description FROM calendar_exceptions ORDER BY date ASC')
      .all();
  }

  return { getExceptionsMap, listExceptions };
}

module.exports = { createCalendarRepository };
