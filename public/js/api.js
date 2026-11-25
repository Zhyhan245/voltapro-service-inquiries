/**
 * Спільний клієнтський модуль взаємодії з REST API та форматування даних.
 */

export function escapeHtml(str) {
  return String(str ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

export async function apiRequest(url, options = {}) {
  const fetchOptions = {
    credentials: 'same-origin',
    headers: {
      'Content-Type': 'application/json',
      ...(options.headers || {})
    },
    ...options
  };

  if (fetchOptions.body && typeof fetchOptions.body !== 'string') {
    fetchOptions.body = JSON.stringify(fetchOptions.body);
  }

  const response = await fetch(url, fetchOptions);
  if (response.status === 204) {
    return null;
  }

  let data = null;
  try {
    data = await response.json();
  } catch {
    data = null;
  }

  if (!response.ok) {
    const err = new Error(data?.message || `Помилка запиту (${response.status})`);
    err.status = response.status;
    err.code = data?.code || 'HTTP_ERROR';
    err.fieldErrors = Array.isArray(data?.fieldErrors) ? data.fieldErrors : [];
    throw err;
  }

  return data;
}

export const STATUS_META = {
  NEW: { label: 'Нова', badgeClass: 'text-bg-primary' },
  IN_PROGRESS: { label: 'В роботі', badgeClass: 'text-bg-warning' },
  COMPLETED: { label: 'Завершена', badgeClass: 'text-bg-success' },
  REJECTED: { label: 'Відхилена', badgeClass: 'text-bg-secondary' }
};

export function renderStatusBadge(status) {
  const meta = STATUS_META[status] || { label: status, badgeClass: 'text-bg-dark' };
  return `<span class="badge ${meta.badgeClass}">${escapeHtml(meta.label)}</span>`;
}

export function renderSlaBadge(slaState, overdueDays = 0, remainingDays = 0) {
  switch (slaState) {
    case 'OVERDUE':
      return `<span class="badge text-bg-danger">Прострочено на ${Number(overdueDays)} роб. дн.</span>`;
    case 'DUE_TODAY':
      return `<span class="badge text-bg-warning">Дедлайн сьогодні</span>`;
    case 'ON_TIME':
      return `<span class="badge text-bg-success">У межах строку (залиш. ${Number(remainingDays)} роб. дн.)</span>`;
    case 'CLOSED_OVERDUE':
      return `<span class="badge text-bg-secondary">Закрито із запізненням (${Number(overdueDays)} роб. дн.)</span>`;
    case 'CLOSED_ON_TIME':
      return `<span class="badge text-bg-light border text-dark">Закрито вчасно</span>`;
    default:
      return `<span class="badge text-bg-light border text-dark">${escapeHtml(slaState)}</span>`;
  }
}

export function formatDateTimeUa(isoString) {
  if (!isoString) return '—';
  const dt = new Date(isoString);
  if (Number.isNaN(dt.getTime())) return escapeHtml(isoString);
  return dt.toLocaleString('uk-UA', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit'
  });
}

export function generateIdempotencyKey() {
  if (window.crypto?.randomUUID) {
    return window.crypto.randomUUID();
  }
  return `req-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}
