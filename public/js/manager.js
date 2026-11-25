import {
  apiRequest,
  escapeHtml,
  formatDateTimeUa,
  renderSlaBadge,
  renderStatusBadge
} from './api.js';

const loginSection = document.getElementById('login-section');
const dashboardSection = document.getElementById('dashboard-section');
const userPanel = document.getElementById('user-panel');
const userBadge = document.getElementById('user-badge');
const logoutBtn = document.getElementById('logout-btn');

const loginForm = document.getElementById('login-form');
const loginEmail = document.getElementById('login-email');
const loginPassword = document.getElementById('login-password');
const loginAlert = document.getElementById('login-alert');
const loginBtn = document.getElementById('login-btn');

const filterForm = document.getElementById('filter-form');
const fltSearch = document.getElementById('flt-search');
const fltStatus = document.getElementById('flt-status');
const fltScope = document.getElementById('flt-scope');
const fltOverdue = document.getElementById('flt-overdue');
const listStatus = document.getElementById('list-status');
const inquiriesTbody = document.getElementById('inquiries-tbody');
const paginationInfo = document.getElementById('pagination-info');
const prevPageBtn = document.getElementById('prev-page-btn');
const nextPageBtn = document.getElementById('next-page-btn');

const slaRefDate = document.getElementById('sla-ref-date');
const slaActiveTotal = document.getElementById('sla-active-total');
const slaOverdueCount = document.getElementById('sla-overdue-count');
const slaDueTodayCount = document.getElementById('sla-due-today-count');
const slaOntimeCount = document.getElementById('sla-ontime-count');
const overdueTbody = document.getElementById('overdue-tbody');
const calendarTbody = document.getElementById('calendar-tbody');
const refreshSlaBtn = document.getElementById('refresh-sla-btn');

const inquiryModalEl = document.getElementById('inquiryModal');
const inquiryModal = new window.bootstrap.Modal(inquiryModalEl);
const modalTitle = document.getElementById('inquiryModalLabel');
const modalSubtitle = document.getElementById('modal-subtitle');
const modalAlert = document.getElementById('modal-alert');
const modalContact = document.getElementById('modal-contact');
const modalSlaInfo = document.getElementById('modal-sla-info');
const modalMessage = document.getElementById('modal-message');
const modalAssignee = document.getElementById('modal-assignee');
const assignBtn = document.getElementById('assign-btn');
const statusForm = document.getElementById('status-form');
const modalNextStatus = document.getElementById('modal-next-status');
const rejectionReasonWrap = document.getElementById('rejection-reason-wrap');
const modalRejectionReason = document.getElementById('modal-rejection-reason');
const modalStatusPublic = document.getElementById('modal-status-public');
const commentForm = document.getElementById('comment-form');
const modalCommentText = document.getElementById('modal-comment-text');
const modalIsInternal = document.getElementById('modal-is-internal');
const modalHistory = document.getElementById('modal-history');

let currentEmployee = null;
let employeesList = [];
let currentPage = 1;
const pageSize = 10;
let activeInquiryId = null;

async function init() {
  try {
    const data = await apiRequest('/api/auth/me');
    setAuthenticated(data.employee);
  } catch {
    setUnauthenticated();
  }
}

function setAuthenticated(employee) {
  currentEmployee = employee;
  loginSection.classList.add('d-none');
  dashboardSection.classList.remove('d-none');
  userPanel.classList.remove('d-none');
  userPanel.classList.add('d-flex');

  const roleLabel = employee.role === 'SENIOR_MANAGER' ? 'Старший менеджер' : 'Менеджер';
  userBadge.textContent = `${employee.fullName} (${roleLabel})`;

  loadEmployees();
  loadInquiries();
  loadSlaReport();
}

function setUnauthenticated() {
  currentEmployee = null;
  dashboardSection.classList.add('d-none');
  userPanel.classList.add('d-none');
  userPanel.classList.remove('d-flex');
  loginSection.classList.remove('d-none');
}

loginForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  loginAlert.classList.add('d-none');
  loginBtn.disabled = true;
  loginBtn.textContent = 'Вхід...';
  try {
    const data = await apiRequest('/api/auth/login', {
      method: 'POST',
      body: {
        email: loginEmail.value.trim(),
        password: loginPassword.value
      }
    });
    loginPassword.value = '';
    setAuthenticated(data.employee);
  } catch (err) {
    loginAlert.textContent = err.message;
    loginAlert.classList.remove('d-none');
  } finally {
    loginBtn.disabled = false;
    loginBtn.textContent = 'Увійти';
  }
});

logoutBtn.addEventListener('click', async () => {
  await apiRequest('/api/auth/logout', { method: 'POST' });
  setUnauthenticated();
});

async function loadEmployees() {
  try {
    const data = await apiRequest('/api/manager/employees');
    employeesList = data.items || [];
  } catch {
    employeesList = [];
  }
}

async function loadInquiries() {
  const params = new URLSearchParams({
    page: String(currentPage),
    pageSize: String(pageSize)
  });
  if (fltSearch.value.trim()) params.set('search', fltSearch.value.trim());
  if (fltStatus.value) params.set('status', fltStatus.value);
  if (fltScope.value) params.set('scope', fltScope.value);
  if (fltOverdue.checked) params.set('overdueOnly', 'true');

  listStatus.textContent = 'Завантаження списку звернень...';
  try {
    const data = await apiRequest(`/api/manager/inquiries?${params.toString()}`);
    renderInquiriesTable(data);
  } catch (err) {
    listStatus.textContent = err.message;
  }
}

function renderInquiriesTable(data) {
  const items = data.items || [];
  const total = data.total || 0;
  const totalPages = Math.max(1, Math.ceil(total / pageSize));

  listStatus.textContent = `Усього знайдено звернень: ${total}`;
  paginationInfo.textContent = `Сторінка ${data.page} з ${totalPages}`;
  prevPageBtn.disabled = data.page <= 1;
  nextPageBtn.disabled = data.page >= totalPages;

  if (items.length === 0) {
    inquiriesTbody.innerHTML = `
      <tr>
        <td colspan="7" class="text-center text-muted py-4">
          Звернень за вказаними фільтрами не знайдено.
        </td>
      </tr>`;
    return;
  }

  inquiriesTbody.innerHTML = items
    .map(
      (item) => `
      <tr>
        <td class="fw-semibold text-nowrap">${escapeHtml(item.ticketNumber)}</td>
        <td>
          <div class="fw-semibold">${escapeHtml(item.subject)}</div>
          <div class="small text-muted">${escapeHtml(item.serviceTitle)}</div>
        </td>
        <td>
          <div>${escapeHtml(item.contactName)}</div>
          <div class="small text-muted">${escapeHtml(item.companyName || item.email)}</div>
        </td>
        <td>${escapeHtml(item.assignedEmployeeName || 'Не призначено')}</td>
        <td>${renderStatusBadge(item.status)}</td>
        <td>
          <div>${renderSlaBadge(item.slaState, item.overdueBusinessDays, item.remainingBusinessDays)}</div>
          <div class="small text-muted">до ${escapeHtml(item.dueDate)}</div>
        </td>
        <td class="text-end">
          <button type="button" class="btn btn-sm btn-outline-primary open-inquiry-btn" data-id="${Number(item.id)}">
            Відкрити
          </button>
        </td>
      </tr>`
    )
    .join('');

  for (const btn of inquiriesTbody.querySelectorAll('.open-inquiry-btn')) {
    btn.addEventListener('click', () => openInquiryModal(Number(btn.dataset.id)));
  }
}

async function openInquiryModal(id) {
  activeInquiryId = id;
  modalAlert.classList.add('d-none');
  try {
    const detail = await apiRequest(`/api/manager/inquiries/${id}`);
    renderInquiryModal(detail);
    inquiryModal.show();
  } catch (err) {
    listStatus.textContent = err.message;
  }
}

function renderInquiryModal(detail) {
  modalTitle.textContent = `${detail.ticketNumber} — ${detail.subject}`;
  modalSubtitle.textContent = `${detail.serviceTitle} (${detail.serviceCategory}) | Створено: ${formatDateTimeUa(detail.createdAt)}`;

  modalContact.innerHTML = `
    <div><strong>Ім’я:</strong> ${escapeHtml(detail.contact.name)}</div>
    <div><strong>Компанія:</strong> ${escapeHtml(detail.contact.companyName || '—')}</div>
    <div><strong>Телефон:</strong> ${escapeHtml(detail.contact.phone)}</div>
    <div><strong>Email:</strong> ${escapeHtml(detail.contact.email)}</div>
  `;

  modalSlaInfo.innerHTML = `
    <div class="mb-1">${renderStatusBadge(detail.status)} ${renderSlaBadge(detail.slaState, detail.overdueBusinessDays, detail.remainingBusinessDays)}</div>
    <div><strong>Норматив SLA:</strong> ${Number(detail.slaBusinessDays)} роб. дн.</div>
    <div><strong>Дедлайн:</strong> ${escapeHtml(detail.dueDate)}</div>
    ${detail.rejectionReason ? `<div class="text-danger mt-1"><strong>Причина відхилення:</strong> ${escapeHtml(detail.rejectionReason)}</div>` : ''}
  `;

  modalMessage.textContent = detail.message;

  modalAssignee.innerHTML = '';
  if (currentEmployee.role === 'SENIOR_MANAGER') {
    const unopt = document.createElement('option');
    unopt.value = '';
    unopt.textContent = '— Не призначено —';
    modalAssignee.appendChild(unopt);
    for (const emp of employeesList) {
      const opt = document.createElement('option');
      opt.value = String(emp.id);
      opt.textContent = `${emp.full_name} (${emp.role === 'SENIOR_MANAGER' ? 'Старший' : 'Менеджер'})`;
      modalAssignee.appendChild(opt);
    }
  } else {
    const opt = document.createElement('option');
    opt.value = String(currentEmployee.id);
    opt.textContent = `${currentEmployee.fullName} (призначити на себе)`;
    modalAssignee.appendChild(opt);
  }
  modalAssignee.value = detail.assignedEmployeeId ? String(detail.assignedEmployeeId) : '';

  modalNextStatus.value = detail.status === 'NEW' ? 'IN_PROGRESS' : 'COMPLETED';
  if (detail.status === 'COMPLETED' || detail.status === 'REJECTED') {
    modalNextStatus.value = 'IN_PROGRESS';
  }
  toggleRejectionWrap();
  modalRejectionReason.value = '';
  modalStatusPublic.value = '';

  modalHistory.innerHTML = (detail.history || [])
    .map((h) => {
      const cls = h.isInternal ? 'note-internal' : 'note-public';
      const tag = h.isInternal
        ? '<span class="badge text-bg-warning">Внутрішня примітка</span>'
        : '<span class="badge text-bg-info text-dark">Публічно</span>';
      return `
        <div class="p-3 rounded ${cls}">
          <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 small mb-1">
            <span><strong>${escapeHtml(h.authorName)}</strong> ${tag}</span>
            <span class="text-muted">${formatDateTimeUa(h.createdAt)}</span>
          </div>
          <div>${escapeHtml(h.comment)}</div>
        </div>`;
    })
    .join('');
}

function showModalAlert(message, type = 'danger') {
  modalAlert.className = `alert alert-${type}`;
  modalAlert.textContent = message;
  modalAlert.classList.remove('d-none');
}

function toggleRejectionWrap() {
  const isRejected = modalNextStatus.value === 'REJECTED';
  rejectionReasonWrap.classList.toggle('d-none', !isRejected);
}

modalNextStatus.addEventListener('change', toggleRejectionWrap);

assignBtn.addEventListener('click', async () => {
  if (!activeInquiryId) return;
  modalAlert.classList.add('d-none');
  try {
    const val = modalAssignee.value ? Number(modalAssignee.value) : null;
    const updated = await apiRequest(`/api/manager/inquiries/${activeInquiryId}/assign`, {
      method: 'PATCH',
      body: { employeeId: val }
    });
    renderInquiryModal(updated);
    showModalAlert('Відповідального працівника оновлено.', 'success');
    loadInquiries();
    loadSlaReport();
  } catch (err) {
    showModalAlert(err.message);
  }
});

statusForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  if (!activeInquiryId) return;
  modalAlert.classList.add('d-none');
  try {
    const updated = await apiRequest(`/api/manager/inquiries/${activeInquiryId}/status`, {
      method: 'PATCH',
      body: {
        status: modalNextStatus.value,
        rejectionReason: modalRejectionReason.value.trim(),
        publicResponse: modalStatusPublic.value.trim()
      }
    });
    renderInquiryModal(updated);
    showModalAlert('Статус звернення успішно оновлено.', 'success');
    loadInquiries();
    loadSlaReport();
  } catch (err) {
    showModalAlert(err.message);
  }
});

commentForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  if (!activeInquiryId) return;
  modalAlert.classList.add('d-none');
  try {
    const updated = await apiRequest(`/api/manager/inquiries/${activeInquiryId}/comments`, {
      method: 'POST',
      body: {
        comment: modalCommentText.value.trim(),
        isInternal: modalIsInternal.checked
      }
    });
    modalCommentText.value = '';
    renderInquiryModal(updated);
    showModalAlert('Запис додано до історії звернення.', 'success');
    loadInquiries();
  } catch (err) {
    showModalAlert(err.message);
  }
});

async function loadSlaReport() {
  try {
    const report = await apiRequest('/api/manager/reports/overdue');
    slaRefDate.textContent = report.referenceDate;
    slaActiveTotal.textContent = String(report.summary.activeTotal);
    slaOverdueCount.textContent = String(report.summary.overdueCount);
    slaDueTodayCount.textContent = String(report.summary.dueTodayCount);
    slaOntimeCount.textContent = String(report.summary.onTimeCount);

    if (report.overdue.length === 0) {
      overdueTbody.innerHTML = `
        <tr><td colspan="7" class="text-center text-muted py-3">Прострочених звернень немає.</td></tr>`;
    } else {
      overdueTbody.innerHTML = report.overdue
        .map(
          (row) => `
          <tr>
            <td class="fw-semibold">${escapeHtml(row.ticketNumber)}</td>
            <td>${escapeHtml(row.subject)}</td>
            <td>${escapeHtml(row.contactName)}</td>
            <td>${escapeHtml(row.assignedEmployeeName)}</td>
            <td>${escapeHtml(row.dueDate)}</td>
            <td><span class="badge text-bg-danger">+${Number(row.overdueBusinessDays)} роб. дн.</span></td>
            <td class="text-end">
              <button type="button" class="btn btn-sm btn-outline-primary open-sla-item" data-id="${Number(row.id)}">Відкрити</button>
            </td>
          </tr>`
        )
        .join('');

      for (const btn of overdueTbody.querySelectorAll('.open-sla-item')) {
        btn.addEventListener('click', () => openInquiryModal(Number(btn.dataset.id)));
      }
    }

    calendarTbody.innerHTML = (report.calendarExceptions || [])
      .map(
        (ex) => `
        <tr>
          <td class="font-monospace">${escapeHtml(ex.date)}</td>
          <td>${ex.is_working_day ? '<span class="badge text-bg-primary">Робочий день</span>' : '<span class="badge text-bg-secondary">Вихідний / свято</span>'}</td>
          <td>${escapeHtml(ex.description)}</td>
        </tr>`
      )
      .join('');
  } catch {
    // Ігноруємо при виході
  }
}

filterForm.addEventListener('submit', (e) => {
  e.preventDefault();
  currentPage = 1;
  loadInquiries();
});

prevPageBtn.addEventListener('click', () => {
  if (currentPage > 1) {
    currentPage -= 1;
    loadInquiries();
  }
});

nextPageBtn.addEventListener('click', () => {
  currentPage += 1;
  loadInquiries();
});

refreshSlaBtn.addEventListener('click', loadSlaReport);

init();
