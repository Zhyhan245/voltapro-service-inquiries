import {
  apiRequest,
  escapeHtml,
  formatDateTimeUa,
  renderSlaBadge,
  renderStatusBadge
} from './api.js';

const lookupForm = document.getElementById('lookup-form');
const ticketInput = document.getElementById('ticket-number');
const codeInput = document.getElementById('access-code');
const lookupBtn = document.getElementById('lookup-btn');
const fillDemoBtn = document.getElementById('fill-demo-btn');
const lookupAlert = document.getElementById('lookup-alert');

const resultCard = document.getElementById('result-card');
const resTicket = document.getElementById('res-ticket');
const resService = document.getElementById('res-service');
const resBadges = document.getElementById('res-badges');
const resSubject = document.getElementById('res-subject');
const resAssignee = document.getElementById('res-assignee');
const resDue = document.getElementById('res-due');
const resCreated = document.getElementById('res-created');
const resRejectionBox = document.getElementById('res-rejection-box');
const resRejection = document.getElementById('res-rejection');
const resPublicBox = document.getElementById('res-public-response-box');
const resPublicResponse = document.getElementById('res-public-response');
const resHistory = document.getElementById('res-history');

async function performLookup(ticketNumber, accessCode) {
  lookupAlert.classList.add('d-none');
  resultCard.classList.add('d-none');
  lookupBtn.disabled = true;
  lookupBtn.textContent = 'Пошук...';

  try {
    const params = new URLSearchParams({
      ticketNumber: ticketNumber.trim().toUpperCase(),
      accessCode: accessCode.trim().toUpperCase()
    });
    const data = await apiRequest(`/api/inquiries/lookup?${params.toString()}`);

    resTicket.textContent = data.ticketNumber;
    resService.textContent = `${data.serviceTitle} (${data.serviceCategory})`;
    resBadges.innerHTML =
      renderStatusBadge(data.status) +
      renderSlaBadge(data.slaState, data.overdueBusinessDays, 0);
    resSubject.textContent = data.subject;
    resAssignee.textContent = data.assignedEmployeeName;
    resDue.textContent = `${data.dueDate} (${data.slaBusinessDays} роб. дн.)`;
    resCreated.textContent = formatDateTimeUa(data.createdAt);

    if (data.status === 'REJECTED' && data.rejectionReason) {
      resRejection.textContent = data.rejectionReason;
      resRejectionBox.classList.remove('d-none');
    } else {
      resRejectionBox.classList.add('d-none');
    }

    if (data.publicResponse) {
      resPublicResponse.textContent = data.publicResponse;
      resPublicBox.classList.remove('d-none');
    } else {
      resPublicBox.classList.add('d-none');
    }

    resHistory.innerHTML = (data.publicHistory || [])
      .map(
        (h) => `
        <div class="p-3 rounded note-public">
          <div class="d-flex justify-content-between small text-muted mb-1">
            <strong>${escapeHtml(h.authorName)}</strong>
            <span>${formatDateTimeUa(h.createdAt)}</span>
          </div>
          <div>${escapeHtml(h.comment)}</div>
        </div>`
      )
      .join('');

    resultCard.classList.remove('d-none');
  } catch (err) {
    lookupAlert.className = 'alert alert-danger mt-3 mb-0';
    lookupAlert.textContent = err.message;
    lookupAlert.classList.remove('d-none');
  } finally {
    lookupBtn.disabled = false;
    lookupBtn.textContent = 'Знайти заявку';
  }
}

lookupForm.addEventListener('submit', (event) => {
  event.preventDefault();
  performLookup(ticketInput.value, codeInput.value);
});

fillDemoBtn.addEventListener('click', () => {
  ticketInput.value = 'VP-2026-0001';
  codeInput.value = 'DEMO01';
  performLookup(ticketInput.value, codeInput.value);
});

const urlParams = new URLSearchParams(window.location.search);
const initialTicket = urlParams.get('ticketNumber');
const initialCode = urlParams.get('accessCode');
if (initialTicket && initialCode) {
  ticketInput.value = initialTicket;
  codeInput.value = initialCode;
  performLookup(initialTicket, initialCode);
}
