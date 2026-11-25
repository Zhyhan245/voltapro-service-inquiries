const express = require('express');

/**
 * Публічні маршрути корпоративного сайту та приймання звернень.
 */
function createPublicRoutes({ db, serviceRepo, calendarRepo, inquiryService }) {
  const router = express.Router();

  router.get('/health', (_req, res) => {
    try {
      db.prepare('SELECT 1 AS ok').get();
      res.json({ status: 'ok', database: 'connected', timestamp: new Date().toISOString() });
    } catch {
      res.status(503).json({ status: 'degraded', database: 'unavailable' });
    }
  });

  router.get('/services', (req, res) => {
    const search = typeof req.query.search === 'string' ? req.query.search.trim() : '';
    const category = typeof req.query.category === 'string' ? req.query.category.trim() : '';
    const items = serviceRepo.list({ search, category, onlyActive: true });
    res.json({ items });
  });

  router.get('/calendar/exceptions', (_req, res) => {
    res.json({ items: calendarRepo.listExceptions() });
  });

  router.post('/inquiries', (req, res, next) => {
    try {
      const result = inquiryService.createPublicInquiry(req.body);
      res.status(result.duplicate ? 200 : 201).json(result);
    } catch (err) {
      next(err);
    }
  });

  router.get('/inquiries/lookup', (req, res, next) => {
    try {
      const result = inquiryService.lookupPublicInquiry(req.query);
      res.json(result);
    } catch (err) {
      next(err);
    }
  });

  return router;
}

module.exports = { createPublicRoutes };
