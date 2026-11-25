const fs = require('node:fs');
const path = require('node:path');
const { performance } = require('node:perf_hooks');
const { createApp } = require('../src/app');
const { seedDatabase } = require('../src/db/seed');

function percentile(sorted, p) {
  if (sorted.length === 0) return 0;
  const idx = Math.min(sorted.length - 1, Math.floor((p / 100) * sorted.length));
  return Number(sorted[idx].toFixed(2));
}

async function runStage(baseUrl, vus, durationMs, cookie) {
  const latencies = [];
  let totalRequests = 0;
  let errorCount = 0;
  const stopAt = performance.now() + durationMs;

  async function worker(workerId) {
    let i = 0;
    while (performance.now() < stopAt) {
      i += 1;
      const t0 = performance.now();
      try {
        const mode = i % 3;
        let res;
        if (mode === 0) {
          res = await fetch(`${baseUrl}/api/services`);
        } else if (mode === 1) {
          res = await fetch(`${baseUrl}/api/inquiries/lookup?ticketNumber=VP-2026-0001&accessCode=DEMO01`);
        } else {
          res = await fetch(`${baseUrl}/api/manager/inquiries?page=1&pageSize=10`, {
            headers: { Cookie: cookie }
          });
        }
        await res.arrayBuffer();
        const dt = performance.now() - t0;
        latencies.push(dt);
        totalRequests += 1;
        if (!res.ok) errorCount += 1;
      } catch {
        errorCount += 1;
        totalRequests += 1;
      }
      // Невелика пауза імітації користувача (20 мс)
      await new Promise((r) => setTimeout(r, 20));
    }
    return workerId;
  }

  const startWall = performance.now();
  const cpuStart = process.cpuUsage();
  await Promise.all(Array.from({ length: vus }, (_, idx) => worker(idx)));
  const elapsedSec = (performance.now() - startWall) / 1000;
  const cpuUsed = process.cpuUsage(cpuStart);
  const memMb = Number((process.memoryUsage().rss / (1024 * 1024)).toFixed(2));

  latencies.sort((a, b) => a - b);

  return {
    virtualUsers: vus,
    durationSec: Number(elapsedSec.toFixed(2)),
    totalRequests,
    rps: Number((totalRequests / elapsedSec).toFixed(2)),
    errorRatePercent: Number(((errorCount / Math.max(1, totalRequests)) * 100).toFixed(3)),
    p50Ms: percentile(latencies, 50),
    p95Ms: percentile(latencies, 95),
    p99Ms: percentile(latencies, 99),
    rssMemoryMb: memMb,
    cpuUserMs: Math.round(cpuUsed.user / 1000)
  };
}

async function main() {
  const { app, db } = createApp({ dbPath: ':memory:', nodeEnv: 'test' });
  seedDatabase(db);

  const server = await new Promise((resolve) => {
    const s = app.listen(0, '127.0.0.1', () => resolve(s));
  });
  const baseUrl = `http://127.0.0.1:${server.address().port}`;

  try {
    const loginRes = await fetch(`${baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        email: 'senior@voltapro.ua',
        password: 'VoltaPro#Manager2026!'
      })
    });
    const cookie = loginRes.headers.get('set-cookie').split(';')[0];

    // Прогрів 1 с
    await runStage(baseUrl, 3, 1000, cookie);

    const stages = [];
    for (const vus of [5, 10, 25]) {
      stages.push(await runStage(baseUrl, vus, 5000, cookie));
    }

    const outDir = path.join(__dirname, '..', 'evidence', 'reports');
    fs.mkdirSync(outDir, { recursive: true });
    const outPath = path.join(outDir, 'load-report.json');
    fs.writeFileSync(
      outPath,
      JSON.stringify({ executedAt: new Date().toISOString(), stages }, null, 2)
    );
    console.log('Навантажувальний звіт збережено у', outPath);
    console.table(stages);
  } finally {
    await new Promise((res) => server.close(res));
    db.close();
  }
}

main();
