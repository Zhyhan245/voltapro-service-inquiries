const express = require('express');

/**
 * Захищені маршрути кабінету менеджера та звітності SLA.
 */
function createManagerRoutes({ requireAuth, employeeRepo, inquiryService }) {
  const router = express.Router();
  router.use(requireAuth);

  router.get('/employees', (_req, res) => {
    res.json({ items: employeeRepo.listActive() });
  });

  router.get('/inquiries', (req, res, next) => {
    try {
      const result = inquiryService.listManagerInquiries(req.employee, req.query);
      res.json(result);
    } catch (err) {
      next(err);
    }
  });

  router.get('/inquiries/:id', (req, res, next) => {
    try {
      const result = inquiryService.getManagerInquiry(req.employee, req.params.id);
      res.json(result);
    } catch (err) {
      next(err);
    }
  });

  router.patch('/inquiries/:id/assign', (req, res, next) => {
    try {
      const result = inquiryService.assignInquiry(
        req.employee,
        req.params.id,
        req.body?.employeeId
      );
      res.json(result);
    } catch (err) {
      next(err);
    }
  });

  router.patch('/inquiries/:id/status', (req, res, next) => {
    try {
      const result = inquiryService.updateInquiryStatus(req.employee, req.params.id, req.body);
      res.json(result);
    } catch (err) {
      next(err);
    }
  });

  router.post('/inquiries/:id/comments', (req, res, next) => {
    try {
      const result = inquiryService.addInquiryComment(req.employee, req.params.id, req.body);
      res.status(201).json(result);
    } catch (err) {
      next(err);
    }
  });

  router.get('/reports/overdue', (req, res, next) => {
    try {
      const result = inquiryService.getOverdueReport(req.employee, req.query.referenceDate);
      res.json(result);
    } catch (err) {
      next(err);
    }
  });

  return router;
}

module.exports = { createManagerRoutes };
