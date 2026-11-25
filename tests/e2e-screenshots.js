const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const { chromium } = require('playwright-core');
const { createApp } = require('../src/app');
const { seedDatabase } = require('../src/db/seed');

const CHROME_PATH = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';

async function main() {
  const outDir = path.join(__dirname, '..', 'evidence', 'screenshots');
  const repDir = path.join(__dirname, '..', 'evidence', 'reports');
  fs.mkdirSync(outDir, { recursive: true });
  fs.mkdirSync(repDir, { recursive: true });

  const { app, db } = createApp({ dbPath: ':memory:', nodeEnv: 'test' });
  seedDatabase(db);

  const server = await new Promise((resolve) => {
    const s = app.listen(0, '127.0.0.1', () => resolve(s));
  });
  const baseUrl = `http://127.0.0.1:${server.address().port}`;

  const browser = await chromium.launch({
    executablePath: CHROME_PATH,
    headless: true
  });

  const checks = [];

  try {
    const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
    const page = await context.newPage();

    // 1. Головна сторінка 1280px + вимірювання навігації
    await page.goto(baseUrl, { waitUntil: 'networkidle' });
    await page.waitForSelector('.service-card');
    const perfHome = await page.evaluate(() => {
      const nav = performance.getEntriesByType('navigation')[0];
      return {
        domContentLoadedMs: Math.round(nav.domContentLoadedEventEnd),
        loadEventMs: Math.round(nav.loadEventEnd),
        hasHorizontalOverflow: document.documentElement.scrollWidth > document.documentElement.clientWidth
      };
    });
    assert.equal(perfHome.hasHorizontalOverflow, false, 'Не повинно бути горизонтального скролу на 1280px');
    await page.screenshot({ path: path.join(outDir, '01-public-desktop-1280.png'), fullPage: true });
    checks.push({ name: 'Головна сторінка (1280px)', passed: true, ...perfHome });

    // 2. Помилка валідації при неповному контакті + збереження введених даних
    await page.selectOption('#inq-service', '1');
    await page.fill('#inq-name', 'Тест');
    await page.fill('#inq-phone', '123');
    await page.fill('#inq-email', 'bad-email');
    await page.fill('#inq-subject', 'Перевірка підключення інвертора');
    await page.fill('#inq-message', 'Потрібен енергоаудит приватного будинку у місті Одеса.');
    await page.click('#submit-inquiry-btn');
    await page.waitForSelector('#form-alert:not(.d-none)');
    const preservedSubject = await page.inputValue('#inq-subject');
    assert.equal(preservedSubject, 'Перевірка підключення інвертора');
    await page.locator('#inquiry-section').scrollIntoViewIfNeeded();
    await page.screenshot({ path: path.join(outDir, '03-public-form-error-validation.png') });
    checks.push({ name: 'Валідація форми зі збереженням полів', passed: true });

    // 3. Успішне подання звернення та блок підтвердження
    await page.fill('#inq-name', 'Марія Гончаренко');
    await page.fill('#inq-phone', '+380671112233');
    await page.fill('#inq-email', 'maria@example.ua');
    await page.click('#submit-inquiry-btn');
    await page.waitForSelector('#confirmation-card:not(.d-none)');
    const ticketNum = await page.textContent('#conf-ticket');
    assert.match(ticketNum, /^VP-\d{4}-\d{4,}$/);
    await page.locator('#inquiry-section').scrollIntoViewIfNeeded();
    await page.screenshot({ path: path.join(outDir, '04-public-inquiry-confirmed.png') });
    checks.push({ name: 'Підтвердження створення звернення', passed: true, ticketNumber: ticketNum });

    // 4. Мобільна адаптація 360px
    await page.setViewportSize({ width: 360, height: 800 });
    await page.goto(baseUrl, { waitUntil: 'networkidle' });
    await page.waitForSelector('.service-card');
    const mobileOverflow = await page.evaluate(
      () => document.documentElement.scrollWidth > document.documentElement.clientWidth
    );
    assert.equal(mobileOverflow, false, 'Не повинно бути горизонтального скролу на 360px');
    await page.screenshot({ path: path.join(outDir, '02-public-mobile-360.png') });
    checks.push({ name: 'Мобільна адаптивність (360px)', passed: true, hasHorizontalOverflow: mobileOverflow });

    // 5. Сторінка перевірки статусу заявки
    await page.setViewportSize({ width: 1280, height: 800 });
    await page.goto(`${baseUrl}/status`, { waitUntil: 'networkidle' });
    await page.click('#fill-demo-btn');
    await page.waitForSelector('#result-card:not(.d-none)');
    const statusBodyText = await page.textContent('#result-card');
    assert.equal(statusBodyText.includes('гарантійний талон №441'), false, 'Внутрішня примітка не має бути видимою клієнту');
    await page.screenshot({ path: path.join(outDir, '05-status-lookup-public.png'), fullPage: true });
    checks.push({ name: 'Публічний пошук заявки без витоку внутрішніх приміток', passed: true });

    // 6. Кабінет менеджера — екран входу
    await page.goto(`${baseUrl}/manager`, { waitUntil: 'networkidle' });
    await page.waitForSelector('#login-section:not(.d-none)');
    await page.screenshot({ path: path.join(outDir, '06-manager-login.png') });
    checks.push({ name: 'Екран входу менеджера', passed: true });

    // 7. Вхід як старший менеджер та робочий список
    await page.fill('#login-email', 'senior@voltapro.ua');
    await page.fill('#login-password', 'VoltaPro#Manager2026!');
    await page.click('#login-btn');
    await page.waitForSelector('#dashboard-section:not(.d-none)');
    await page.waitForSelector('#inquiries-tbody tr');
    await page.screenshot({ path: path.join(outDir, '07-manager-list-1280.png'), fullPage: true });
    checks.push({ name: 'Робочий список менеджера (1280px)', passed: true });

    // 8. Картка звернення (модальне вікно з історією та внутрішньою приміткою)
    await page.click('.open-inquiry-btn');
    await page.waitForSelector('#inquiryModal.show');
    await page.waitForTimeout(250);
    await page.screenshot({ path: path.join(outDir, '08-manager-modal-detail.png') });
    await page.click('#inquiryModal .btn-close');
    await page.waitForTimeout(250);
    checks.push({ name: 'Картка звернення з внутрішніми примітками', passed: true });

    // 9. Вкладка звіту про прострочені заявки (SLA)
    await page.click('#tab-sla');
    await page.waitForSelector('#pane-sla.active');
    await page.waitForTimeout(200);
    await page.screenshot({ path: path.join(outDir, '09-manager-sla-report.png'), fullPage: true });
    checks.push({ name: 'Звіт про прострочені заявки (SLA)', passed: true });

    // 10. Планшетна адаптація 768px для кабінету менеджера
    await page.setViewportSize({ width: 768, height: 900 });
    await page.click('#tab-inquiries');
    await page.waitForTimeout(200);
    const tabletOverflow = await page.evaluate(
      () => document.documentElement.scrollWidth > document.documentElement.clientWidth
    );
    assert.equal(tabletOverflow, false, 'Не повинно бути горизонтального скролу на 768px');
    await page.screenshot({ path: path.join(outDir, '10-manager-tablet-768.png') });
    checks.push({ name: 'Планшетна адаптивність кабінету (768px)', passed: true, hasHorizontalOverflow: tabletOverflow });

    fs.writeFileSync(
      path.join(repDir, 'e2e-ui-report.json'),
      JSON.stringify({ executedAt: new Date().toISOString(), browser: 'Google Chrome (headless)', checks }, null, 2)
    );

    console.log(`Усі ${checks.length} E2E та UI перевірок пройдено! Скріншоти збережено у ${outDir}`);
  } finally {
    await browser.close();
    await new Promise((res) => server.close(res));
    db.close();
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
