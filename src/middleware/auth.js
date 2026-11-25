const { DomainError } = require('../domain/inquiryRules');

/**
 * Middleware автентифікації через HttpOnly cookie-сесію.
 */
function createAuthMiddleware({ authService, cookieName }) {
  function attachSession(req, _res, next) {
    const token = req.cookies?.[cookieName];
    req.employee = token ? authService.authenticate(token) : null;
    req.sessionToken = token || null;
    next();
  }

  function requireAuth(req, _res, next) {
    if (!req.employee) {
      return next(
        new DomainError(
          'UNAUTHORIZED',
          401,
          'Необхідна автентифікація для доступу до кабінету менеджера.'
        )
      );
    }
    return next();
  }

  return { attachSession, requireAuth };
}

module.exports = { createAuthMiddleware };
