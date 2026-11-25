const express = require('express');

/**
 * Маршрути автентифікації працівників компанії.
 */
function createAuthRoutes({ authService, cookieName, sessionTtlHours, isProd }) {
  const router = express.Router();

  router.post('/login', (req, res, next) => {
    try {
      const { token, expiresAt, employee } = authService.login({
        email: req.body?.email,
        password: req.body?.password
      });

      res.cookie(cookieName, token, {
        httpOnly: true,
        sameSite: 'lax',
        secure: isProd,
        maxAge: sessionTtlHours * 3600 * 1000,
        path: '/'
      });

      res.json({ employee, expiresAt });
    } catch (err) {
      next(err);
    }
  });

  router.post('/logout', (req, res) => {
    authService.logout(req.sessionToken);
    res.clearCookie(cookieName, { path: '/' });
    res.status(204).end();
  });

  router.get('/me', (req, res) => {
    if (!req.employee) {
      return res.status(401).json({
        code: 'UNAUTHORIZED',
        message: 'Сесія відсутня або завершилася.'
      });
    }
    return res.json({ employee: req.employee });
  });

  return router;
}

module.exports = { createAuthRoutes };
