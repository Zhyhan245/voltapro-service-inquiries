/**
 * Доменні правила переходів статусів та розмежування прав доступу до звернень.
 */

const STATUSES = Object.freeze({
  NEW: 'NEW',
  IN_PROGRESS: 'IN_PROGRESS',
  COMPLETED: 'COMPLETED',
  REJECTED: 'REJECTED'
});

const ALLOWED_TRANSITIONS = Object.freeze({
  [STATUSES.NEW]: [STATUSES.IN_PROGRESS, STATUSES.REJECTED],
  [STATUSES.IN_PROGRESS]: [STATUSES.COMPLETED, STATUSES.REJECTED],
  [STATUSES.COMPLETED]: [STATUSES.IN_PROGRESS], // Лише для SENIOR_MANAGER
  [STATUSES.REJECTED]: [STATUSES.IN_PROGRESS]   // Лише для SENIOR_MANAGER
});

const TERMINAL_STATUSES = new Set([STATUSES.COMPLETED, STATUSES.REJECTED]);

class DomainError extends Error {
  constructor(code, status, message, fieldErrors = []) {
    super(message);
    this.name = 'DomainError';
    this.code = code;
    this.status = status;
    this.fieldErrors = fieldErrors;
  }
}

/**
 * Перевіряє допустимість переходу статусу з урахуванням ролі працівника та причини відхилення.
 */
function validateStatusTransition({ currentStatus, nextStatus, employeeRole, rejectionReason }) {
  if (!Object.values(STATUSES).includes(nextStatus)) {
    throw new DomainError('INVALID_STATUS', 400, 'Неприпустимий цільовий статус звернення.', [
      { field: 'status', message: 'Невідомий статус звернення.' }
    ]);
  }

  if (currentStatus === nextStatus) {
    throw new DomainError('SAME_STATUS', 409, 'Звернення вже перебуває у вибраному статусі.');
  }

  const allowedTargets = ALLOWED_TRANSITIONS[currentStatus] || [];
  if (!allowedTargets.includes(nextStatus)) {
    throw new DomainError(
      'TRANSITION_NOT_ALLOWED',
      409,
      `Перехід зі статусу «${currentStatus}» у статус «${nextStatus}» заборонено.`
    );
  }

  if (TERMINAL_STATUSES.has(currentStatus) && employeeRole !== 'SENIOR_MANAGER') {
    throw new DomainError(
      'REOPEN_FORBIDDEN',
      403,
      'Повернути завершене або відхилене звернення в роботу може лише старший менеджер.'
    );
  }

  const trimmedReason = typeof rejectionReason === 'string' ? rejectionReason.trim() : '';
  if (nextStatus === STATUSES.REJECTED && trimmedReason.length < 5) {
    throw new DomainError(
      'REJECTION_REASON_REQUIRED',
      400,
      'При відхиленні звернення обов’язково вкажіть причину (мінімум 5 символів).',
      [{ field: 'rejectionReason', message: 'Вкажіть причину відхилення (мін. 5 символів).' }]
    );
  }

  return trimmedReason || null;
}

/**
 * Перевіряє, чи має працівник право переглядати/редагувати конкретне звернення.
 * Старший менеджер бачить усі заявки. Звичайний менеджер — лише нерозподілені або призначені йому.
 */
function assertEmployeeCanAccessInquiry(employee, inquiry) {
  if (!employee) {
    throw new DomainError('UNAUTHORIZED', 401, 'Потрібна автентифікація.');
  }
  if (employee.role === 'SENIOR_MANAGER') {
    return;
  }
  if (inquiry.assigned_employee_id !== null && inquiry.assigned_employee_id !== employee.id) {
    throw new DomainError(
      'INQUIRY_FORBIDDEN',
      403,
      'Доступ заборонено: звернення закріплене за іншим відповідальним менеджером.'
    );
  }
}

module.exports = {
  STATUSES,
  ALLOWED_TRANSITIONS,
  DomainError,
  validateStatusTransition,
  assertEmployeeCanAccessInquiry
};
