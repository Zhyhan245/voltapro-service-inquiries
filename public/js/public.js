import { apiRequest, escapeHtml, generateIdempotencyKey } from './api.js';

const searchInput = document.getElementById('service-search');
const categorySelect = document.getElementById('service-category');
const servicesStatus = document.getElementById('services-status');
const servicesGrid = document.getElementById('services-grid');
const serviceSelect = document.getElementById('inq-service');

const inquiryForm = document.getElementById('inquiry-form');
const submitBtn = document.getElementById('submit-inquiry-btn');
const formAlert = document.getElementById('form-alert');

const confirmationCard = document.getElementById('confirmation-card');
const confTicket = document.getElementById('conf-ticket');
const confCode = document.getElementById('conf-code');
const confService = document.getElementById('conf-service');
const confSla = document.getElementById('conf-sla');
const confDue = document.getElementById('conf-due');
const confStatusLink = document.getElementById('conf-status-link');

let allServices = [];
let currentIdempotencyKey = generateIdempotencyKey();
let submitting = false;

async function loadServices() {
  const search = searchInput.value.trim();
  const category = categorySelect.value;
  const params = new URLSearchParams();
  if (search) params.set('search', search);
  if (category) params.set('category', category);

  servicesStatus.textContent = 'Завантаження каталогу послуг...';
  try {
    const data = await apiRequest(`/api/services?${params.toString()}`);
    const items = data.items || [];
    if (!search && !category) {
      allServices = items;
      populateServiceSelect(allServices);
    }
    renderServices(items);
  } catch (err) {
    servicesStatus.textContent = '';
    servicesGrid.innerHTML = `<div class="col-12"><div class="alert alert-danger" role="alert">${escapeHtml(err.message)}</div></div>`;
  }
}

function populateServiceSelect(items) {
  const previous = serviceSelect.value;
  serviceSelect.innerHTML = '<option value="">Оберіть послугу...</option>';
  for (const item of items) {
    const opt = document.createElement('option');
    opt.value = String(item.id);
    opt.textContent = `${item.title} (${item.category}, SLA: ${item.sla_business_days} роб. дн.)`;
    serviceSelect.appendChild(opt);
  }
  if (previous) serviceSelect.value = previous;
}

function renderServices(items) {
  if (items.length === 0) {
    servicesStatus.textContent = 'За вашим запитом послуг не знайдено.';
    servicesGrid.innerHTML = `
      <div class="col-12">
        <div class="alert alert-light border text-center py-4">
          Послуг за вибраними критеріями не знайдено. Спробуйте змінити пошуковий запит або категорію.
        </div>
      </div>`;
    return;
  }

  servicesStatus.textContent = `Знайдено послуг: ${items.length}`;
  servicesGrid.innerHTML = items
    .map(
      (s) => `
      <div class="col-md-6 col-lg-3">
        <article class="card h-100 service-card">
          <div class="card-body d-flex flex-column">
            <div class="d-flex justify-content-between align-items-start gap-2 mb-2">
              <span class="badge text-bg-light border text-dark">${escapeHtml(s.category)}</span>
              <span class="badge text-bg-info text-dark">SLA: ${Number(s.sla_business_days)} роб. дн.</span>
            </div>
            <h3 class="h6 fw-bold mb-2">${escapeHtml(s.title)}</h3>
            <p class="small text-muted flex-grow-1 mb-3">${escapeHtml(s.short_description)}</p>
            <div class="d-flex justify-content-between align-items-center pt-2 border-top">
              <span class="fw-bold text-primary">від ${Number(s.base_price_uah).toLocaleString('uk-UA')} грн</span>
              <button type="button" class="btn btn-sm btn-outline-primary select-service-btn" data-service-id="${Number(s.id)}">
                Обрати
              </button>
            </div>
          </div>
        </article>
      </div>`
    )
    .join('');

  for (const btn of servicesGrid.querySelectorAll('.select-service-btn')) {
    btn.addEventListener('click', () => {
      serviceSelect.value = btn.dataset.serviceId;
      document.getElementById('inquiry-section').scrollIntoView({ behavior: 'smooth' });
      serviceSelect.focus();
    });
  }
}

function clearFieldErrors() {
  formAlert.classList.add('d-none');
  formAlert.textContent = '';
  for (const el of inquiryForm.querySelectorAll('.is-invalid')) {
    el.classList.remove('is-invalid');
  }
  for (const fb of inquiryForm.querySelectorAll('[data-error-for]')) {
    fb.textContent = '';
  }
}

function showFieldErrors(fieldErrors, generalMessage) {
  if (generalMessage) {
    formAlert.className = 'alert alert-danger';
    formAlert.textContent = generalMessage;
    formAlert.classList.remove('d-none');
  }
  for (const fe of fieldErrors) {
    const input = inquiryForm.querySelector(`[name="${CSS.escape(fe.field)}"]`);
    const fb = inquiryForm.querySelector(`[data-error-for="${CSS.escape(fe.field)}"]`);
    if (input) input.classList.add('is-invalid');
    if (fb) fb.textContent = fe.message;
  }
}

inquiryForm.addEventListener('submit', async (event) => {
  event.preventDefault();
  if (submitting) return;

  clearFieldErrors();
  const payload = {
    serviceId: Number(inquiryForm.serviceId.value),
    contactName: inquiryForm.contactName.value.trim(),
    companyName: inquiryForm.companyName.value.trim(),
    phone: inquiryForm.phone.value.trim(),
    email: inquiryForm.email.value.trim(),
    subject: inquiryForm.subject.value.trim(),
    message: inquiryForm.message.value.trim(),
    idempotencyKey: currentIdempotencyKey
  };

  submitting = true;
  submitBtn.disabled = true;
  submitBtn.textContent = 'Надсилання...';

  try {
    const res = await apiRequest('/api/inquiries', {
      method: 'POST',
      body: payload
    });

    confTicket.textContent = res.ticketNumber;
    confCode.textContent = res.accessCode;
    confService.textContent = res.serviceTitle;
    confSla.textContent = `${res.slaBusinessDays} роб. дн.`;
    confDue.textContent = res.dueDate;
    confStatusLink.href = `/status?ticketNumber=${encodeURIComponent(res.ticketNumber)}&accessCode=${encodeURIComponent(res.accessCode)}`;
    confirmationCard.classList.remove('d-none');

    formAlert.className = 'alert alert-success';
    formAlert.textContent = `Звернення ${res.ticketNumber} успішно зареєстровано.`;
    formAlert.classList.remove('d-none');

    inquiryForm.reset();
    currentIdempotencyKey = generateIdempotencyKey();
  } catch (err) {
    showFieldErrors(err.fieldErrors || [], err.message);
  } finally {
    submitting = false;
    submitBtn.disabled = false;
    submitBtn.textContent = 'Надіслати звернення';
  }
});

let searchTimer = null;
searchInput.addEventListener('input', () => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(loadServices, 200);
});
categorySelect.addEventListener('change', loadServices);

loadServices();
