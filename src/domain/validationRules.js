const { DomainError } = require('./inquiryRules');

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
const PHONE_RE = /^\+?[\d\s()-]{9,18}$/;

function cleanText(value) {
  return typeof value === 'string' ? value.trim() : '';
}

/**
 * Валідація вхідної форми створення звернення від клієнта.
 */
function validateCreateInquiryInput(body) {
  const fieldErrors = [];

  const serviceId = Number(body?.serviceId);
  if (!Number.isInteger(serviceId) || serviceId <= 0) {
    fieldErrors.push({ field: 'serviceId', message: 'Оберіть послугу зі списку.' });
  }

  const contactName = cleanText(body?.contactName);
  if (contactName.length < 2 || contactName.length > 100) {
    fieldErrors.push({
      field: 'contactName',
      message: 'Ім’я контактної особи має містити від 2 до 100 символів.'
    });
  }

  const companyName = cleanText(body?.companyName);
  if (companyName.length > 120) {
    fieldErrors.push({
      field: 'companyName',
      message: 'Назва компанії не може перевищувати 120 символів.'
    });
  }

  const phone = cleanText(body?.phone);
  const digitsCount = phone.replace(/\D/g, '').length;
  if (!PHONE_RE.test(phone) || digitsCount < 9 || digitsCount > 15) {
    fieldErrors.push({
      field: 'phone',
      message: 'Вкажіть коректний номер телефону (від 9 до 15 цифр).'
    });
  }

  const email = cleanText(body?.email).toLowerCase();
  if (!EMAIL_RE.test(email) || email.length > 120) {
    fieldErrors.push({
      field: 'email',
      message: 'Вкажіть коректну електронну пошту.'
    });
  }

  const subject = cleanText(body?.subject);
  if (subject.length < 4 || subject.length > 160) {
    fieldErrors.push({
      field: 'subject',
      message: 'Тема звернення має містити від 4 до 160 символів.'
    });
  }

  const message = cleanText(body?.message);
  if (message.length < 10 || message.length > 2000) {
    fieldErrors.push({
      field: 'message',
      message: 'Опис звернення має містити від 10 до 2000 символів.'
    });
  }

  const idempotencyKey = cleanText(body?.idempotencyKey);

  if (fieldErrors.length > 0) {
    throw new DomainError(
      'VALIDATION_ERROR',
      400,
      'Перевірте коректність заповнення форми звернення.',
      fieldErrors
    );
  }

  return {
    serviceId,
    contactName,
    companyName: companyName || null,
    phone,
    email,
    subject,
    message,
    idempotencyKey: idempotencyKey || null
  };
}

/**
 * Валідація параметрів пошуку заявки клієнтом за номером та кодом доступу.
 */
function validatePublicLookupQuery(query) {
  const ticketNumber = cleanText(query?.ticketNumber).toUpperCase();
  const accessCode = cleanText(query?.accessCode).toUpperCase();
  const fieldErrors = [];

  if (!/^VP-\d{4}-\d{4,}$/.test(ticketNumber)) {
    fieldErrors.push({
      field: 'ticketNumber',
      message: 'Формат номера заявки: VP-РІК-НОМЕР (наприклад, VP-2026-0001).'
    });
  }
  if (!/^[A-Z0-9]{6}$/.test(accessCode)) {
    fieldErrors.push({
      field: 'accessCode',
      message: 'Код доступу має складатися із 6 символів.'
    });
  }

  if (fieldErrors.length > 0) {
    throw new DomainError(
      'VALIDATION_ERROR',
      400,
      'Вкажіть коректний номер заявки та 6-символьний код доступу.',
      fieldErrors
    );
  }

  return { ticketNumber, accessCode };
}

module.exports = {
  validateCreateInquiryInput,
  validatePublicLookupQuery,
  cleanText
};
