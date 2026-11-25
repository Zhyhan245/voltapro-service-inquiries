const { createApp } = require('./app');
const { logEvent } = require('./utils/logger');

const { app, db, config } = createApp();

const server = app.listen(config.port, () => {
  logEvent('info', 'server', `Сервер «ВольтаПро Сервіс» запущено на http://localhost:${config.port}`, {
    port: config.port,
    nodeEnv: config.nodeEnv
  });
});

function shutdown(signal) {
  logEvent('info', 'server', `Отримано сигнал ${signal}, завершення роботи...`);
  server.close(() => {
    db.close();
    process.exit(0);
  });
}

process.on('SIGINT', () => shutdown('SIGINT'));
process.on('SIGTERM', () => shutdown('SIGTERM'));
