/**
 * Репозиторій працівників та серверних сесій.
 */
function createEmployeeRepository(db) {
  function findByEmail(email) {
    return db
      .prepare(
        `SELECT id, full_name, email, password_hash, role, department, is_active, created_at
         FROM employees
         WHERE LOWER(email) = LOWER(?)`
      )
      .get(email);
  }

  function findById(id) {
    return db
      .prepare(
        `SELECT id, full_name, email, role, department, is_active, created_at
         FROM employees
         WHERE id = ?`
      )
      .get(id);
  }

  function listActive() {
    return db
      .prepare(
        `SELECT id, full_name, email, role, department
         FROM employees
         WHERE is_active = 1
         ORDER BY role DESC, full_name ASC`
      )
      .all();
  }

  function createSession({ token, employeeId, expiresAt, createdAt }) {
    db.prepare(
      `INSERT INTO sessions (token, employee_id, expires_at, created_at)
       VALUES (?, ?, ?, ?)`
    ).run(token, employeeId, expiresAt, createdAt);
  }

  function findSessionWithEmployee(token) {
    return db
      .prepare(
        `SELECT s.token, s.expires_at,
                e.id AS employee_id, e.full_name, e.email, e.role, e.department, e.is_active
         FROM sessions s
         JOIN employees e ON e.id = s.employee_id
         WHERE s.token = ?`
      )
      .get(token);
  }

  function deleteSession(token) {
    db.prepare('DELETE FROM sessions WHERE token = ?').run(token);
  }

  function purgeExpiredSessions(nowIso) {
    db.prepare('DELETE FROM sessions WHERE expires_at <= ?').run(nowIso);
  }

  return {
    findByEmail,
    findById,
    listActive,
    createSession,
    findSessionWithEmployee,
    deleteSession,
    purgeExpiredSessions
  };
}

module.exports = { createEmployeeRepository };
