const { verifyPassword, generateSessionToken } = require('../utils/security');
const { DomainError } = require('../domain/inquiryRules');

/**
 * Прикладний сервіс автентифікації працівників компанії.
 */
function createAuthService({ employeeRepo, sessionTtlHours }) {
  function login({ email, password }) {
    const cleanEmail = typeof email === 'string' ? email.trim().toLowerCase() : '';
    const rawPassword = typeof password === 'string' ? password : '';

    if (!cleanEmail || !rawPassword) {
      throw new DomainError(
        'VALIDATION_ERROR',
        400,
        'Вкажіть електронну пошту та пароль.',
        [
          ...(!cleanEmail ? [{ field: 'email', message: 'Обов’язкове поле.' }] : []),
          ...(!rawPassword ? [{ field: 'password', message: 'Обов’язкове поле.' }] : [])
        ]
      );
    }

    const employee = employeeRepo.findByEmail(cleanEmail);
    if (!employee || !employee.is_active || !verifyPassword(rawPassword, employee.password_hash)) {
      throw new DomainError(
        'INVALID_CREDENTIALS',
        401,
        'Неправильна електронна пошта або пароль.'
      );
    }

    const now = new Date();
    const expires = new Date(now.getTime() + sessionTtlHours * 3600 * 1000);
    const token = generateSessionToken();

    employeeRepo.purgeExpiredSessions(now.toISOString());
    employeeRepo.createSession({
      token,
      employeeId: employee.id,
      expiresAt: expires.toISOString(),
      createdAt: now.toISOString()
    });

    return {
      token,
      expiresAt: expires.toISOString(),
      employee: toEmployeeDto(employee)
    };
  }

  function authenticate(token) {
    if (!token || typeof token !== 'string') return null;
    const row = employeeRepo.findSessionWithEmployee(token);
    if (!row) return null;

    if (!row.is_active || row.expires_at <= new Date().toISOString()) {
      employeeRepo.deleteSession(token);
      return null;
    }

    return {
      id: row.employee_id,
      fullName: row.full_name,
      email: row.email,
      role: row.role,
      department: row.department
    };
  }

  function logout(token) {
    if (token) {
      employeeRepo.deleteSession(token);
    }
  }

  function toEmployeeDto(row) {
    return {
      id: row.id,
      fullName: row.full_name,
      email: row.email,
      role: row.role,
      department: row.department
    };
  }

  return { login, authenticate, logout };
}

module.exports = { createAuthService };
