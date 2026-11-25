const { DomainError } = require('../domain/inquiryRules');
const { logEvent } = require('../utils/logger');

/**
 * Централізований обробник помилок у форматі Problem Details (RFC 9457).
 */
function errorHandler(err, req, res, _next) {
  if (err instanceof DomainError) {
    return res.status(err.status).json({
      type: `https://voltapro.local/errors/${err.code.toLowerCase()}`,
      title: err.message,
      status: err.status,
      code: err.code,
      message: err.message,
      fieldErrors: err.fieldErrors || [],
      requestId: req.requestId
    });
  }

  // Некоректний JSON у тілі запиту
  if (err?.type === 'entity.parse.failed') {
    return res.status(400).json({
      type: 'https://voltapro.local/errors/invalid_json',
      title: 'Некоректний формат JSON у тілі запиту.',
      status: 400,
      code: 'INVALID_JSON',
      message: 'Некоректний формат JSON у тілі запиту.',
      fieldErrors: [],
      requestId: req.requestId
    });
  }

  logEvent('error', 'http', 'Неочікувана серверна помилка', {
    requestId: req.requestId,
    method: req.method,
    path: req.path,
    errorName: err?.name || 'Error'
  });

  return res.status(500).json({
    type: 'https://voltapro.local/errors/internal_error',
    title: 'Внутрішня помилка сервера.',
    status: 500,
    code: 'INTERNAL_ERROR',
    message: 'Сталася внутрішня помилка сервера. Спробуйте пізніше.',
    fieldErrors: [],
    requestId: req.requestId
  });
}

module.exports = { errorHandler };
