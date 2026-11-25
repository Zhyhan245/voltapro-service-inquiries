const test = require('node:test');
const assert = require('node:assert/strict');

const {
  isWorkingDay,
  addBusinessDays,
  countBusinessDaysBetween,
  evaluateSlaState
} = require('../src/domain/calendarRules');
const {
  STATUSES,
  DomainError,
  validateStatusTransition,
  assertEmployeeCanAccessInquiry
} = require('../src/domain/inquiryRules');
const { validateCreateInquiryInput } = require('../src/domain/validationRules');
const { hashPassword, verifyPassword } = require('../src/utils/security');

test('Календар: вихідні (сб/нд) пропускаються, а святкові винятки враховуються', () => {
  const exceptions = new Map([
    ['2026-12-25', false], // п'ятниця, святковий неробочий день
    ['2026-12-12', true]   // субота, перенесений робочий день
  ]);

  assert.equal(isWorkingDay('2026-12-25', exceptions), false);
  assert.equal(isWorkingDay('2026-12-12', exceptions), true);
  assert.equal(isWorkingDay('2026-12-13', exceptions), false);

  // 2026-12-24 (чт) + 2 робочих дні з пропуском 25 (свято), 26 (сб), 27 (нд) -> 2026-12-29 (вт)
  const due = addBusinessDays('2026-12-24', 2, exceptions);
  assert.equal(due, '2026-12-29');
});

test('SLA: визначення станів ON_TIME, DUE_TODAY, OVERDUE та CLOSED', () => {
  const exceptions = new Map();
  const dueFriday = '2026-11-27';

  const onTime = evaluateSlaState({ status: 'IN_PROGRESS', due_date: dueFriday }, '2026-11-25', exceptions);
  assert.equal(onTime.state, 'ON_TIME');
  assert.equal(onTime.remainingBusinessDays, 2);

  const dueToday = evaluateSlaState({ status: 'NEW', due_date: dueFriday }, '2026-11-27', exceptions);
  assert.equal(dueToday.state, 'DUE_TODAY');

  // Пн 2026-11-30 та Вт 2026-12-01 -> +2 робочих дні прострочення
  const overdue = evaluateSlaState({ status: 'IN_PROGRESS', due_date: dueFriday }, '2026-12-01', exceptions);
  assert.equal(overdue.state, 'OVERDUE');
  assert.equal(overdue.overdueBusinessDays, 2);
  assert.equal(countBusinessDaysBetween(dueFriday, '2026-12-01', exceptions), 2);
});

test('Правила статусів: обов’язкова причина відхилення та заборона повернення завершеної заявки для MANAGER', () => {
  assert.throws(
    () =>
      validateStatusTransition({
        currentStatus: STATUSES.IN_PROGRESS,
        nextStatus: STATUSES.REJECTED,
        employeeRole: 'MANAGER',
        rejectionReason: '   '
      }),
    (err) => err instanceof DomainError && err.code === 'REJECTION_REASON_REQUIRED' && err.status === 400
  );

  assert.throws(
    () =>
      validateStatusTransition({
        currentStatus: STATUSES.COMPLETED,
        nextStatus: STATUSES.IN_PROGRESS,
        employeeRole: 'MANAGER'
      }),
    (err) => err instanceof DomainError && err.code === 'REOPEN_FORBIDDEN' && err.status === 403
  );

  const ok = validateStatusTransition({
    currentStatus: STATUSES.COMPLETED,
    nextStatus: STATUSES.IN_PROGRESS,
    employeeRole: 'SENIOR_MANAGER'
  });
  assert.equal(ok, null);
});

test('Правила доступу: менеджер не може відкрити заявку іншого менеджера', () => {
  const manager1 = { id: 2, role: 'MANAGER' };
  assert.throws(
    () => assertEmployeeCanAccessInquiry(manager1, { assigned_employee_id: 3 }),
    (err) => err instanceof DomainError && err.status === 403
  );
  assert.doesNotThrow(() =>
    assertEmployeeCanAccessInquiry({ id: 1, role: 'SENIOR_MANAGER' }, { assigned_employee_id: 3 })
  );
});

test('Валідація форми: відхилення неповного контакту', () => {
  assert.throws(
    () =>
      validateCreateInquiryInput({
        serviceId: 1,
        contactName: 'А',
        phone: '123',
        email: 'bad-email',
        subject: 'Ок',
        message: 'Коротко'
      }),
    (err) => err instanceof DomainError && err.status === 400 && err.fieldErrors.length >= 4
  );
});

test('Безпека: scrypt хешування та верифікація пароля', () => {
  const hash = hashPassword('VoltaPro#Manager2026!');
  assert.equal(verifyPassword('VoltaPro#Manager2026!', hash), true);
  assert.equal(verifyPassword('WrongPassword12345', hash), false);
});
