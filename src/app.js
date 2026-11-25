const path = require('node:path');
const express = require('express');
const cookieParser = require('cookie-parser');
const helmet = require('helmet');

const { loadConfig } = require('./config');
const { openDatabase } = require('./db/connection');
const { runMigrations } = require('./db/migrate');
const { createRequestId, logEvent } = require('./utils/logger');

const { createServiceRepository } = require('./repositories/serviceRepository');
const { createEmployeeRepository } = require('./repositories/employeeRepository');
const { createCalendarRepository } = require('./repositories/calendarRepository');
const { createInquiryRepository } = require('./repositories/inquiryRepository');

const { createAuthService } = require('./services/authService');
const { createInquiryService } = require('./services/inquiryService');
const { createAuthMiddleware } = require('./middleware/auth');
const { errorHandler } = require('./middleware/errorHandler');

const { createPublicRoutes } = require('./routes/publicRoutes');
const { createAuthRoutes } = require('./routes/authRoutes');
const { createManagerRoutes } = require('./routes/managerRoutes');

/**
 * Створює налаштований екземпляр Express-застосунку та підключення до БД.
 */
function createApp(configOverrides = {}) {
  const config = loadConfig(configOverrides);
  const db = configOverrides.db || openDatabase(config.dbPath);
  runMigrations(db);

  const serviceRepo = createServiceRepository(db);
  const employeeRepo = createEmployeeRepository(db);
  const calendarRepo = createCalendarRepository(db);
  const inquiryRepo = createInquiryRepository(db);

  const authService = createAuthService({
    employeeRepo,
    sessionTtlHours: config.sessionTtlHours
  });
  const inquiryService = createInquiryService({
    inquiryRepo,
    serviceRepo,
    employeeRepo,
    calendarRepo
  });

  const { attachSession, requireAuth } = createAuthMiddleware({
    authService,
    cookieName: config.cookieName
  });

  const app = express();
  app.disable('x-powered-by');
  app.use(
    helmet({
      contentSecurityPolicy: {
        directives: {
          defaultSrc: ["'self'"],
          scriptSrc: ["'self'", 'https://cdn.jsdelivr.net'],
          styleSrc: ["'self'", 'https://cdn.jsdelivr.net', "'unsafe-inline'"],
          fontSrc: ["'self'", 'https://cdn.jsdelivr.net'],
          imgSrc: ["'self'", 'data:'],
          connectSrc: ["'self'"]
        }
      }
    })
  );
  app.use(express.json({ limit: '64kb' }));
  app.use(cookieParser());

  app.use((req, res, next) => {
    req.requestId = createRequestId();
    res.setHeader('X-Request-Id', req.requestId);
    const start = Date.now();
    res.on('finish', () => {
      if (config.nodeEnv !== 'test' && req.path.startsWith('/api')) {
        logEvent('info', 'http', `${req.method} ${req.path}`, {
          requestId: req.requestId,
          status: res.statusCode,
          durationMs: Date.now() - start
        });
      }
    });
    next();
  });

  app.use(attachSession);

  app.use(
    '/api',
    createPublicRoutes({ db, serviceRepo, calendarRepo, inquiryService })
  );
  app.use(
    '/api/auth',
    createAuthRoutes({
      authService,
      cookieName: config.cookieName,
      sessionTtlHours: config.sessionTtlHours,
      isProd: config.nodeEnv === 'production'
    })
  );
  app.use(
    '/api/manager',
    createManagerRoutes({ requireAuth, employeeRepo, inquiryService })
  );

  // Статичні файли публічного сайту, локальний Bootstrap 5 та кабінет менеджера
  const publicDir = path.join(__dirname, '..', 'public');
  const bootstrapDir = path.join(__dirname, '..', 'node_modules', 'bootstrap', 'dist');
  app.use('/vendor/bootstrap', express.static(bootstrapDir));
  app.use(express.static(publicDir));

  // Маршрути інтерфейсу
  app.get('/manager', (_req, res) => {
    res.sendFile(path.join(publicDir, 'manager.html'));
  });
  app.get('/status', (_req, res) => {
    res.sendFile(path.join(publicDir, 'status.html'));
  });

  // 404 для неіснуючих API-маршрутів
  app.use('/api', (_req, res) => {
    res.status(404).json({
      code: 'ENDPOINT_NOT_FOUND',
      message: 'Запитаний маршрут API не знайдено.'
    });
  });

  app.use(errorHandler);

  return { app, db, config, services: { authService, inquiryService } };
}

module.exports = { createApp };
