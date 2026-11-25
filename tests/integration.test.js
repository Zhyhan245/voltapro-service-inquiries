const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');

const { createApp } = require('../src/app');
const { seedDatabase } = require('../src/db/seed');

function startTestServer() {
  const { app, db } = createApp({ dbPath: ':memory:', nodeEnv: 'test' });
  seedDatabase(db);
  return new Promise((resolve) => {
    const server = app.listen(0, '127.0.0.1', () => {
      const addr = server.address();
      resolve({
        baseUrl: `http://127.0.0.1:${addr.port}`,
        db,
        close: () =>
          new Promise((res) => {
            server.close(() => {
              db.close();
              res();
            });
          })
      });
    });
  });
}

async function httpJson(url, options = {}) {
  const res = await fetch(url, {
    ...options,
    headers: {
      'Content-Type': 'application/json',
      ...(options.headers || {})
    }
  });
  let body = null;
  if (res.status !== 204) {
    body = await res.json().catch(() => null);
  }
  return { status: res.status, headers: res.headers, body };
}

test('Інтеграційні сценарії Варіанту 5: створення, ідемпотентність, приватність нотаток, права і звіт SLA', async (t) => {
  const srv = await startTestServer();
  t.after(() => srv.close());

  // 1. Каталог послуг і фільтрація (включно з кириличним пошуком без урахування регістру)
  const servicesRes = await httpJson(`${srv.baseUrl}/api/services?category=${encodeURIComponent('Аудит')}`);
  assert.equal(servicesRes.status, 200);
  assert.ok(servicesRes.body.items.length >= 2);
  const serviceId = servicesRes.body.items[0].id;

  const cyrillicSearchRes = await httpJson(`${srv.baseUrl}/api/services?search=${encodeURIComponent('монтаж')}`);
  assert.equal(cyrillicSearchRes.status, 200);
  assert.ok(cyrillicSearchRes.body.items.length >= 2);

  // 2. Відхилення неповного контакту
  const badCreate = await httpJson(`${srv.baseUrl}/api/inquiries`, {
    method: 'POST',
    body: JSON.stringify({
      serviceId,
      contactName: 'Я',
      phone: '123',
      email: 'invalid',
      subject: 'Т',
      message: 'Мало'
    })
  });
  assert.equal(badCreate.status, 400);
  assert.equal(badCreate.body.code, 'VALIDATION_ERROR');

  // 3. Коректне створення + повторне надсилання з тим самим idempotencyKey (без дубліката)
  const idempKey = 'test-idempotency-key-001';
  const createPayload = {
    serviceId,
    contactName: 'Олег Сидоренко',
    companyName: 'ТОВ «ТестЕнерго»',
    phone: '+380671234567',
    email: 'oleg@testenergo.ua',
    subject: 'Аудит офісної мережі на вул. Канатній',
    message: 'Прошу провести аудит ввідного щита та оцінити потребу в резервуванні.',
    idempotencyKey: idempKey
  };

  const firstCreate = await httpJson(`${srv.baseUrl}/api/inquiries`, {
    method: 'POST',
    body: JSON.stringify(createPayload)
  });
  assert.equal(firstCreate.status, 201);
  assert.match(firstCreate.body.ticketNumber, /^VP-\d{4}-\d{4,}$/);
  assert.equal(firstCreate.body.duplicate, false);

  const duplicateCreate = await httpJson(`${srv.baseUrl}/api/inquiries`, {
    method: 'POST',
    body: JSON.stringify(createPayload)
  });
  assert.equal(duplicateCreate.status, 200);
  assert.equal(duplicateCreate.body.duplicate, true);
  assert.equal(duplicateCreate.body.ticketNumber, firstCreate.body.ticketNumber);

  // 4. Неавторизований доступ до кабінету менеджера -> 401
  const unauthRes = await httpJson(`${srv.baseUrl}/api/manager/inquiries`);
  assert.equal(unauthRes.status, 401);

  // 5. Вхід manager1 і робота із заявкою
  const loginM1 = await httpJson(`${srv.baseUrl}/api/auth/login`, {
    method: 'POST',
    body: JSON.stringify({
      email: 'manager1@voltapro.ua',
      password: 'VoltaPro#Manager2026!'
    })
  });
  assert.equal(loginM1.status, 200);
  const cookieM1 = loginM1.headers.get('set-cookie').split(';')[0];

  // Знаходимо створену заявку в списку
  const listM1 = await httpJson(
    `${srv.baseUrl}/api/manager/inquiries?search=${encodeURIComponent(firstCreate.body.ticketNumber)}`,
    { headers: { Cookie: cookieM1 } }
  );
  assert.equal(listM1.status, 200);
  assert.equal(listM1.body.items.length, 1);
  const createdId = listM1.body.items[0].id;

  // Призначаємо на себе (manager1)
  const assignRes = await httpJson(`${srv.baseUrl}/api/manager/inquiries/${createdId}/assign`, {
    method: 'PATCH',
    headers: { Cookie: cookieM1 },
    body: JSON.stringify({ employeeId: loginM1.body.employee.id })
  });
  assert.equal(assignRes.status, 200);

  // Додаємо внутрішню примітку та публічний коментар
  const secretInternalText = 'ВНУТРІШНЯ ПРИМІТКА: клієнт просив знижку 15%, не підтверджено.';
  await httpJson(`${srv.baseUrl}/api/manager/inquiries/${createdId}/comments`, {
    method: 'POST',
    headers: { Cookie: cookieM1 },
    body: JSON.stringify({ comment: secretInternalText, isInternal: true })
  });

  const publicText = 'Інженер зв’яжеться з вами завтра о 10:00 для уточнення доступу.';
  await httpJson(`${srv.baseUrl}/api/manager/inquiries/${createdId}/comments`, {
    method: 'POST',
    headers: { Cookie: cookieM1 },
    body: JSON.stringify({ comment: publicText, isInternal: false })
  });

  // 6. Перевірка публічного пошуку: внутрішня примітка НЕ повинна потрапити клієнту!
  const publicLookup = await httpJson(
    `${srv.baseUrl}/api/inquiries/lookup?ticketNumber=${encodeURIComponent(firstCreate.body.ticketNumber)}&accessCode=${encodeURIComponent(firstCreate.body.accessCode)}`
  );
  assert.equal(publicLookup.status, 200);
  assert.equal(publicLookup.body.publicResponse, publicText);
  const serializedPublic = JSON.stringify(publicLookup.body);
  assert.equal(serializedPublic.includes(secretInternalText), false);

  // Неправильний accessCode -> 404
  const badLookup = await httpJson(
    `${srv.baseUrl}/api/inquiries/lookup?ticketNumber=${encodeURIComponent(firstCreate.body.ticketNumber)}&accessCode=ZZZZ99`
  );
  assert.equal(badLookup.status, 404);

  // 7. Вхід manager2 -> спроба відкрити чужу заявку manager1 -> 403
  const loginM2 = await httpJson(`${srv.baseUrl}/api/auth/login`, {
    method: 'POST',
    body: JSON.stringify({
      email: 'manager2@voltapro.ua',
      password: 'VoltaPro#Manager2026!'
    })
  });
  const cookieM2 = loginM2.headers.get('set-cookie').split(';')[0];

  const foreignAccess = await httpJson(`${srv.baseUrl}/api/manager/inquiries/${createdId}`, {
    headers: { Cookie: cookieM2 }
  });
  assert.equal(foreignAccess.status, 403);

  // 8. Переходи статусів, відмова без причини (400), завершення й спроба повернення без права (403)
  const rejectNoReason = await httpJson(`${srv.baseUrl}/api/manager/inquiries/${createdId}/status`, {
    method: 'PATCH',
    headers: { Cookie: cookieM1 },
    body: JSON.stringify({ status: 'REJECTED', rejectionReason: '' })
  });
  assert.equal(rejectNoReason.status, 400);

  const toInProgress = await httpJson(`${srv.baseUrl}/api/manager/inquiries/${createdId}/status`, {
    method: 'PATCH',
    headers: { Cookie: cookieM1 },
    body: JSON.stringify({ status: 'IN_PROGRESS' })
  });
  assert.equal(toInProgress.status, 200);

  const toCompleted = await httpJson(`${srv.baseUrl}/api/manager/inquiries/${createdId}/status`, {
    method: 'PATCH',
    headers: { Cookie: cookieM1 },
    body: JSON.stringify({ status: 'COMPLETED', publicResponse: 'Аудит завершено.' })
  });
  assert.equal(toCompleted.status, 200);

  const reopenByManager = await httpJson(`${srv.baseUrl}/api/manager/inquiries/${createdId}/status`, {
    method: 'PATCH',
    headers: { Cookie: cookieM1 },
    body: JSON.stringify({ status: 'IN_PROGRESS' })
  });
  assert.equal(reopenByManager.status, 403);

  // Старший менеджер може повернути в роботу
  const loginSenior = await httpJson(`${srv.baseUrl}/api/auth/login`, {
    method: 'POST',
    body: JSON.stringify({
      email: 'senior@voltapro.ua',
      password: 'VoltaPro#Manager2026!'
    })
  });
  const cookieSenior = loginSenior.headers.get('set-cookie').split(';')[0];

  const reopenBySenior = await httpJson(`${srv.baseUrl}/api/manager/inquiries/${createdId}/status`, {
    method: 'PATCH',
    headers: { Cookie: cookieSenior },
    body: JSON.stringify({ status: 'IN_PROGRESS', publicResponse: 'Повернуто на додаткову перевірку.' })
  });
  assert.equal(reopenBySenior.status, 200);

  // 9. Звіт про прострочені заявки SLA
  const reportRes = await httpJson(`${srv.baseUrl}/api/manager/reports/overdue`, {
    headers: { Cookie: cookieSenior }
  });
  assert.equal(reportRes.status, 200);
  assert.ok(reportRes.body.summary.overdueCount >= 1);
});
