const fs = require('node:fs');
const path = require('node:path');
const { performance } = require('node:perf_hooks');
const { openDatabase, runInTransaction } = require('../src/db/connection');
const { runMigrations } = require('../src/db/migrate');
const { seedDatabase } = require('../src/db/seed');
const { createInquiryRepository } = require('../src/repositories/inquiryRepository');
const { createCalendarRepository } = require('../src/repositories/calendarRepository');
const { createEmployeeRepository } = require('../src/repositories/employeeRepository');
const { createServiceRepository } = require('../src/repositories/serviceRepository');
const { createInquiryService } = require('../src/services/inquiryService');

function populateSyntheticInquiries(db, targetCount) {
  const existing = Number(db.prepare('SELECT COUNT(*) AS c FROM inquiries').get().c);
  const toInsert = targetCount - existing;
  if (toInsert <= 0) return;

  const serviceIds = db.prepare('SELECT id, sla_business_days FROM services').all();
  const employeeIds = db.prepare('SELECT id FROM employees').all().map((r) => r.id);
  const statuses = ['NEW', 'IN_PROGRESS', 'COMPLETED', 'REJECTED'];

  const insertContact = db.prepare(
    'INSERT INTO client_contacts (contact_name, company_name, phone, email, created_at) VALUES (?, ?, ?, ?, ?)'
  );
  const insertInq = db.prepare(
    `INSERT INTO inquiries (
      ticket_number, access_code, service_id, contact_id, assigned_employee_id,
      subject, message, status, rejection_reason, public_response, idempotency_key,
      sla_business_days, due_date, created_at, updated_at, closed_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, NULL, NULL, ?, ?, ?, ?, ?, ?)`
  );

  runInTransaction(db, () => {
    for (let i = 0; i < toInsert; i += 1) {
      const idx = existing + i + 1;
      const srv = serviceIds[idx % serviceIds.length];
      const status = statuses[idx % statuses.length];
      const assigned = idx % 5 === 0 ? null : employeeIds[idx % employeeIds.length];
      const dayShift = (idx % 20) - 10;
      const dt = new Date(Date.UTC(2026, 10, 25 + dayShift));
      const due = dt.toISOString().slice(0, 10);
      const nowIso = new Date().toISOString();

      const cRes = insertContact.run(
        `Клієнт Тестовий ${idx}`,
        idx % 2 === 0 ? `ТОВ «Енерго-${idx}»` : null,
        `+38067${String(idx).padStart(7, '0').slice(-7)}`,
        `client${idx}@example.ua`,
        nowIso
      );
      const ticket = `VP-2026-${String(idx + 10000).padStart(5, '0')}`;
      insertInq.run(
        ticket,
        `V${String(idx).padStart(5, '0').slice(-5)}`,
        srv.id,
        Number(cRes.lastInsertRowid),
        assigned,
        `Тестове звернення обсягу #${idx}`,
        'Синтетичний опис звернення для перевірки продуктивності вибірок і пошуку.',
        status,
        `vol-${idx}`,
        srv.sla_business_days,
        due,
        nowIso,
        nowIso,
        status === 'COMPLETED' || status === 'REJECTED' ? nowIso : null
      );
    }
  });
}

function measureMs(fn, iterations = 10) {
  const times = [];
  for (let i = 0; i < iterations; i += 1) {
    const t0 = performance.now();
    fn();
    times.push(performance.now() - t0);
  }
  times.sort((a, b) => a - b);
  const avg = times.reduce((s, x) => s + x, 0) / times.length;
  return {
    avgMs: Number(avg.toFixed(3)),
    p95Ms: Number(times[Math.min(times.length - 1, Math.floor(times.length * 0.95))].toFixed(3))
  };
}

function runVolumeExperiment() {
  const sizes = [100, 1000, 10000];
  const results = [];

  const tmpDbPath = path.join(__dirname, '..', 'data', 'volume-bench.db');
  if (fs.existsSync(tmpDbPath)) fs.unlinkSync(tmpDbPath);

  const db = openDatabase(tmpDbPath);
  try {
    runMigrations(db);
    seedDatabase(db);

    const inquiryRepo = createInquiryRepository(db);
    const serviceRepo = createServiceRepository(db);
    const employeeRepo = createEmployeeRepository(db);
    const calendarRepo = createCalendarRepository(db);
    const inquiryService = createInquiryService({
      inquiryRepo,
      serviceRepo,
      employeeRepo,
      calendarRepo
    });
    const senior = { id: 1, role: 'SENIOR_MANAGER' };

    for (const size of sizes) {
      populateSyntheticInquiries(db, size);

      const listMetric = measureMs(() =>
        inquiryService.listManagerInquiries(senior, { status: 'IN_PROGRESS', page: 1, pageSize: 15 })
      );
      const searchMetric = measureMs(() =>
        inquiryService.listManagerInquiries(senior, { search: 'Клієнт Тестовий 50', page: 1, pageSize: 15 })
      );
      const slaReportMetric = measureMs(() => inquiryService.getOverdueReport(senior));

      const explainPlan = db
        .prepare(
          "EXPLAIN QUERY PLAN SELECT id, ticket_number FROM inquiries WHERE status IN ('NEW', 'IN_PROGRESS') AND due_date < '2026-11-25'"
        )
        .all()
        .map((r) => r.detail);

      results.push({
        recordsCount: size,
        listFiltered: listMetric,
        searchByText: searchMetric,
        overdueReport: slaReportMetric,
        queryPlan: explainPlan
      });
    }
  } finally {
    db.close();
    if (fs.existsSync(tmpDbPath)) fs.unlinkSync(tmpDbPath);
  }

  const outDir = path.join(__dirname, '..', 'evidence', 'reports');
  fs.mkdirSync(outDir, { recursive: true });
  const outPath = path.join(outDir, 'volume-report.json');
  fs.writeFileSync(
    outPath,
    JSON.stringify({ executedAt: new Date().toISOString(), results }, null, 2)
  );
  console.log('Звіт тестування обсягом збережено у', outPath);
  console.table(
    results.map((r) => ({
      Записів: r.recordsCount,
      'Список (avg мс)': r.listFiltered.avgMs,
      'Пошук (avg мс)': r.searchByText.avgMs,
      'Звіт SLA (avg мс)': r.overdueReport.avgMs
    }))
  );
}

runVolumeExperiment();
