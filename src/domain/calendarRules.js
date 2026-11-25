/**
 * Правила виробничого календаря та розрахунку строків опрацювання (SLA).
 * Робочі дні — понеділок–п'ятниця (UTC), якщо інше не задано в calendar_exceptions.
 */

function parseIsoDate(dateStr) {
  if (typeof dateStr !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    throw new Error(`Некоректний формат дати (очікується YYYY-MM-DD): ${dateStr}`);
  }
  const [y, m, d] = dateStr.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  if (
    dt.getUTCFullYear() !== y ||
    dt.getUTCMonth() !== m - 1 ||
    dt.getUTCDate() !== d
  ) {
    throw new Error(`Неіснуюча календарна дата: ${dateStr}`);
  }
  return dt;
}

function formatIsoDate(dateObj) {
  return dateObj.toISOString().slice(0, 10);
}

function toIsoDateOnly(input) {
  if (!input) return formatIsoDate(new Date());
  if (input instanceof Date) return formatIsoDate(input);
  if (/^\d{4}-\d{2}-\d{2}/.test(input)) return input.slice(0, 10);
  return formatIsoDate(new Date(input));
}

function addCalendarDays(dateStr, days) {
  const dt = parseIsoDate(dateStr);
  dt.setUTCDate(dt.getUTCDate() + days);
  return formatIsoDate(dt);
}

/**
 * Визначає, чи є дата робочим днем з урахуванням вихідних (сб/нд) та винятків календаря.
 */
function isWorkingDay(dateStr, exceptionsMap = new Map()) {
  if (exceptionsMap.has(dateStr)) {
    return Boolean(exceptionsMap.get(dateStr));
  }
  const day = parseIsoDate(dateStr).getUTCDay();
  return day !== 0 && day !== 6;
}

/**
 * Додає N робочих днів до дати реєстрації звернення.
 */
function addBusinessDays(startDateStr, businessDays, exceptionsMap = new Map()) {
  if (!Number.isInteger(businessDays) || businessDays < 1) {
    throw new Error('Кількість робочих днів SLA має бути додатним цілим числом.');
  }
  let cursor = toIsoDateOnly(startDateStr);
  let remaining = businessDays;
  let guard = 0;

  while (remaining > 0) {
    cursor = addCalendarDays(cursor, 1);
    if (isWorkingDay(cursor, exceptionsMap)) {
      remaining -= 1;
    }
    guard += 1;
    if (guard > 3660) {
      throw new Error('Перевищено ліміт ітерацій календаря (понад 10 років).');
    }
  }
  return cursor;
}

/**
 * Рахує кількість робочих днів між двома датами (від startExclusive до endInclusive).
 */
function countBusinessDaysBetween(startExclusive, endInclusive, exceptionsMap = new Map()) {
  const start = toIsoDateOnly(startExclusive);
  const end = toIsoDateOnly(endInclusive);
  if (end <= start) return 0;

  let cursor = start;
  let count = 0;
  let guard = 0;
  while (cursor < end) {
    cursor = addCalendarDays(cursor, 1);
    if (isWorkingDay(cursor, exceptionsMap)) {
      count += 1;
    }
    guard += 1;
    if (guard > 3660) {
      throw new Error('Перевищено ліміт ітерацій підрахунку робочих днів.');
    }
  }
  return count;
}

/**
 * Оцінює поточний стан SLA звернення на вказану дату.
 */
function evaluateSlaState(inquiry, referenceDateStr = toIsoDateOnly(), exceptionsMap = new Map()) {
  const today = toIsoDateOnly(referenceDateStr);
  const dueDate = toIsoDateOnly(inquiry.due_date ?? inquiry.dueDate);

  if (inquiry.status === 'COMPLETED' || inquiry.status === 'REJECTED') {
    const closedDate = toIsoDateOnly(inquiry.closed_at ?? inquiry.closedAt ?? today);
    const overdueBusinessDays =
      closedDate > dueDate ? countBusinessDaysBetween(dueDate, closedDate, exceptionsMap) : 0;
    return {
      state: overdueBusinessDays > 0 ? 'CLOSED_OVERDUE' : 'CLOSED_ON_TIME',
      overdueBusinessDays,
      remainingBusinessDays: 0
    };
  }

  if (today > dueDate) {
    const overdueBusinessDays = countBusinessDaysBetween(dueDate, today, exceptionsMap);
    return {
      state: overdueBusinessDays > 0 ? 'OVERDUE' : 'DUE_TODAY',
      overdueBusinessDays,
      remainingBusinessDays: 0
    };
  }

  if (today === dueDate) {
    return {
      state: 'DUE_TODAY',
      overdueBusinessDays: 0,
      remainingBusinessDays: 0
    };
  }

  return {
    state: 'ON_TIME',
    overdueBusinessDays: 0,
    remainingBusinessDays: countBusinessDaysBetween(today, dueDate, exceptionsMap)
  };
}

module.exports = {
  isWorkingDay,
  addBusinessDays,
  countBusinessDaysBetween,
  evaluateSlaState,
  toIsoDateOnly
};
