const { loadConfig } = require('../config');
const { openDatabase, runInTransaction } = require('./connection');
const { runMigrations } = require('./migrate');
const { hashPassword } = require('../utils/security');
const { addBusinessDays } = require('../domain/calendarRules');

/**
 * Ідемпотентне заповнення БД демонстраційними синтетичними даними.
 * Повторний запуск не створює дублікатів.
 */
function seedDatabase(db) {
  runMigrations(db);

  const existingServices = db.prepare('SELECT COUNT(*) AS cnt FROM services').get();
  if (Number(existingServices.cnt) > 0) {
    return { seeded: false, reason: 'База даних вже містить демонстраційні дані.' };
  }

  const now = new Date();
  const nowIso = now.toISOString();
  const todayStr = nowIso.slice(0, 10);

  const demoPasswordHash = hashPassword('VoltaPro#Manager2026!');

  return runInTransaction(db, () => {
    // 1. Послуги підприємства (8 записів, 4 категорії)
    const servicesData = [
      {
        slug: 'solar-audit-home',
        title: 'Енергоаудит приватного будинку',
        category: 'Аудит',
        short_description: 'Обстеження електромережі, аналіз профілю споживання та підбір потужності інвертора.',
        full_description: 'Виїзд інженера, заміри пікових навантажень, перевірка заземлення та підготовка технічного звіту з рекомендаціями щодо резервного живлення.',
        base_price_uah: 2500,
        sla_business_days: 2
      },
      {
        slug: 'solar-audit-business',
        title: 'Енергоаудит комерційного об’єкта',
        category: 'Аудит',
        short_description: 'Інструментальний аудит електропостачання офісу, складу чи виробництва.',
        full_description: 'Комплексний аналіз трифазної мережі, гармонійних спотворень, резервування критичних ліній та оцінка окупності СЕС.',
        base_price_uah: 6800,
        sla_business_days: 3
      },
      {
        slug: 'design-backup-5kw',
        title: 'Проєктування системи резервного живлення до 10 кВт',
        category: 'Проєктування',
        short_description: 'Розроблення однолінійної схеми, специфікації обладнання та плану щита АВР.',
        full_description: 'Повний пакет робочої документації для безпечного підключення гібридного інвертора та LiFePO4 АКБ згідно з ПУЕ.',
        base_price_uah: 4200,
        sla_business_days: 3
      },
      {
        slug: 'design-solar-commercial',
        title: 'Проєктування мережевої СЕС для бізнесу (від 30 кВт)',
        category: 'Проєктування',
        short_description: '3D-моделювання затінення, розрахунок генерації та проєктна документація.',
        full_description: 'Включає інженерні розрахунки вітрового навантаження покрівлі, підбір стрінгів та підготовку схеми приєднання.',
        base_price_uah: 14500,
        sla_business_days: 5
      },
      {
        slug: 'install-hybrid-inverter',
        title: 'Монтаж гібридного інвертора та АКБ LiFePO4',
        category: 'Монтаж',
        short_description: 'Встановлення, обв’язка захисною автоматикою, пусконалагодження та інструктаж.',
        full_description: 'Монтаж обладнання «під ключ» із налаштуванням пріоритетів заряду, тестуванням перемикання на резерв та гарантією на роботи.',
        base_price_uah: 9500,
        sla_business_days: 4
      },
      {
        slug: 'install-solar-panels',
        title: 'Монтаж сонячних панелей на покрівлю',
        category: 'Монтаж',
        short_description: 'Кріплення алюмінієвого профілю, укладання ФЕМ, прокладання солар-кабелю та грозозахист.',
        full_description: 'Роботи виконуються сертифікованою бригадою з дотриманням герметизації покрівлі та тестуванням вольт-амперних характеристик.',
        base_price_uah: 12000,
        sla_business_days: 5
      },
      {
        slug: 'maintenance-diagnostics',
        title: 'Планове ТО інвертора та батарейного блоку',
        category: 'Обслуговування',
        short_description: 'Тепловізійний контроль контактів, оновлення прошивки та перевірка балансування АКБ.',
        full_description: 'Профілактичне обслуговування для збереження гарантії виробника та запобігання аварійним відключенням у піковий сезон.',
        base_price_uah: 1900,
        sla_business_days: 2
      },
      {
        slug: 'maintenance-emergency',
        title: 'Термінова діагностика та відновлення живлення',
        category: 'Обслуговування',
        short_description: 'Оперативний виїзд інженера при помилках інвертора, спрацюванні захисту чи втраті зв’язку.',
        full_description: 'Пріоритетний розбір звернення, зчитування журналу аварій, заміна пошкоджених модулів захисту.',
        base_price_uah: 3200,
        sla_business_days: 1
      }
    ];

    const insertService = db.prepare(
      `INSERT INTO services (
        slug, title, category, short_description, full_description,
        base_price_uah, sla_business_days, is_active, created_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?)`
    );

    const serviceIds = [];
    for (const s of servicesData) {
      const res = insertService.run(
        s.slug,
        s.title,
        s.category,
        s.short_description,
        s.full_description,
        s.base_price_uah,
        s.sla_business_days,
        nowIso
      );
      serviceIds.push({ id: Number(res.lastInsertRowid), ...s });
    }

    // 2. Працівники (3 облікові записи)
    const employeesData = [
      {
        full_name: 'Ковальчук Олена Сергіївна',
        email: 'senior@voltapro.ua',
        role: 'SENIOR_MANAGER',
        department: 'Керівництво клієнтського сервісу'
      },
      {
        full_name: 'Бондаренко Андрій Вікторович',
        email: 'manager1@voltapro.ua',
        role: 'MANAGER',
        department: 'Відділ проєктування та аудиту'
      },
      {
        full_name: 'Ткаченко Марія Ігорівна',
        email: 'manager2@voltapro.ua',
        role: 'MANAGER',
        department: 'Відділ монтажу та сервісу'
      }
    ];

    const insertEmployee = db.prepare(
      `INSERT INTO employees (
        full_name, email, password_hash, role, department, is_active, created_at
      ) VALUES (?, ?, ?, ?, ?, 1, ?)`
    );

    const employeeIds = [];
    for (const emp of employeesData) {
      const res = insertEmployee.run(
        emp.full_name,
        emp.email,
        demoPasswordHash,
        emp.role,
        emp.department,
        nowIso
      );
      employeeIds.push(Number(res.lastInsertRowid));
    }

    // 3. Виробничий календар (святкові/вихідні дні)
    const calendarData = [
      ['2026-12-25', 0, 'Різдво Христове (святковий неробочий день)'],
      ['2027-01-01', 0, 'Новий рік (святковий неробочий день)'],
      ['2027-01-07', 0, 'Святковий неробочий день'],
      ['2026-12-12', 1, 'Робоча субота (перенесення графіка для завершення річних об’єктів)']
    ];

    const insertCal = db.prepare(
      'INSERT INTO calendar_exceptions (date, is_working_day, description) VALUES (?, ?, ?)'
    );
    const exceptionsMap = new Map();
    for (const [dt, isWork, desc] of calendarData) {
      insertCal.run(dt, isWork, desc);
      exceptionsMap.set(dt, isWork === 1);
    }

    // 4. Демонстраційні звернення (18 записів у різних статусах та станах SLA)
    function shiftDays(baseIsoDate, deltaDays) {
      const dt = new Date(`${baseIsoDate}T10:00:00.000Z`);
      dt.setUTCDate(dt.getUTCDate() + deltaDays);
      return dt.toISOString();
    }

    const demoInquiries = [
      {
        code: 'DEMO01',
        serviceIdx: 7,
        assignedIdx: 1,
        contact: ['Савченко Ігор Петрович', 'ТОВ «Одеса-Логістик»', '+380671112233', 'savchenko@odesalog.ua'],
        subject: 'Помилка F56 на інверторі Deye 12 кВт на складі',
        message: 'Після ранкового стрибка напруги інвертор перейшов у байпас і не заряджає батарею.',
        status: 'IN_PROGRESS',
        daysOffsetCreated: -5,
        dueDateOverride: shiftDays(todayStr, -2).slice(0, 10), // Прострочено
        publicResp: 'Інженер виїхав на об’єкт, проводимо діагностику плати BMS.',
        internalNote: 'Клієнт просив терміновий виїзд до 16:00, перевірити гарантійний талон №441.'
      },
      {
        code: 'DEMO02',
        serviceIdx: 0,
        assignedIdx: null,
        contact: ['Мельник Наталія Олегівна', null, '+380502223344', 'n.melnyk@example.ua'],
        subject: 'Аудит електромережі таунхауса перед купівлею АКБ',
        message: 'Потрібно оцінити можливість підключення резервного живлення на 8 кВт у щитовій.',
        status: 'NEW',
        daysOffsetCreated: -4,
        dueDateOverride: shiftDays(todayStr, -1).slice(0, 10), // Прострочено
        publicResp: null,
        internalNote: 'Необхідно терміново призначити менеджера відділу аудиту.'
      },
      {
        code: 'DEMO03',
        serviceIdx: 4,
        assignedIdx: 2,
        contact: ['Коваленко Дмитро Андрійович', 'Аптека «Здоров’я+»', '+380933334455', 'kovalenko@apteka.ua'],
        subject: 'Монтаж інвертора та двох модулів АКБ для холодильного обладнання',
        message: 'Необхідне безперебійне живлення холодильних камер на 6 годин автономної роботи.',
        status: 'IN_PROGRESS',
        daysOffsetCreated: -3,
        dueDateOverride: todayStr, // Дедлайн сьогодні
        publicResp: 'Кошторис погоджено, бригада сформована на завтрашній ранок.',
        internalNote: 'Закуплено автоматику ETI зі складу №2.'
      },
      {
        code: 'DEMO04',
        serviceIdx: 2,
        assignedIdx: 1,
        contact: ['Гриценко Олександр Іванович', null, '+380664445566', 'grytsenko@example.ua'],
        subject: 'Проєкт щита АВР для приватного будинку в Чорноморську',
        message: 'Маю генератор та гібридний інвертор, потрібна коректна схема розв’язки вводів.',
        status: 'COMPLETED',
        daysOffsetCreated: -7,
        dueDateOverride: shiftDays(todayStr, -2).slice(0, 10),
        publicResp: 'Робочий проєкт та однолінійну схему надіслано на електронну пошту.',
        internalNote: 'Оплату підтверджено бухгалтерією.'
      },
      {
        code: 'DEMO05',
        serviceIdx: 5,
        assignedIdx: 2,
        contact: ['Лисенко Сергій Миколайович', 'ФОП Лисенко С.М.', '+380975556677', 'lysenko.sm@example.ua'],
        subject: 'Монтаж 16 панелей на шиферний дах старого зразка',
        message: 'Хочемо встановити 16 сонячних панелей по 550 Вт на покрівлю 1978 року.',
        status: 'REJECTED',
        rejectionReason: 'Несуча здатність дерев’яних крокв не відповідає вимогам безпеки без реконструкції даху.',
        daysOffsetCreated: -6,
        dueDateOverride: shiftDays(todayStr, -1).slice(0, 10),
        publicResp: 'На жаль, монтаж неможливий без попереднього підсилення кроквяної системи.',
        internalNote: 'Фото тріщин крокв додано в архів інженера.'
      },
      {
        code: 'DEMO06',
        serviceIdx: 1,
        assignedIdx: 1,
        contact: ['Павленко Вікторія Сергіївна', 'БЦ «Південний»', '+380676667788', 'pavlenko@bc-south.ua'],
        subject: 'Інструментальний аудит трифазного вводу 120 кВт',
        message: 'Плануємо встановлення промислової СЕС для зменшення денного споживання.',
        status: 'NEW',
        daysOffsetCreated: 0,
        dueDateOverride: null,
        publicResp: null,
        internalNote: 'Потенційно великий корпоративний клієнт.'
      }
    ];

    // Додамо ще 12 стандартних заявок для пагінації та фільтрації
    const extraClients = [
      ['Кравченко Олена', 'Стоматологія «Дента»', '+380501234501', 'denta@example.ua'],
      ['Шевчук Тарас', null, '+380671234502', 'shevchuk@example.ua'],
      ['Поліщук Ірина', 'Кав’ярня «Зерно»', '+380931234503', 'zerno@example.ua'],
      ['Бойко Максим', 'СТО «Турбо»', '+380661234504', 'turbo@example.ua'],
      ['Романюк Анна', null, '+380971234505', 'romanyuk@example.ua'],
      ['Гаврилюк Вадим', 'ТОВ «АгроПлюс»', '+380671234506', 'agroplus@example.ua'],
      ['Дяченко Юлія', null, '+380501234507', 'dyachenko@example.ua'],
      ['Семенюк Богдан', 'Пекарня «Хлібний дім»', '+380931234508', 'bread@example.ua'],
      ['Левченко Віктор', null, '+380661234509', 'levchenko@example.ua'],
      ['Мороз Катерина', 'Салон «Естетика»', '+380971234510', 'moroz@example.ua'],
      ['Остапчук Денис', 'Логістика «Південь»', '+380671234511', 'ostapchuk@example.ua'],
      ['Федоренко Світлана', null, '+380501234512', 'fedorenko@example.ua']
    ];

    const extraStatuses = ['NEW', 'IN_PROGRESS', 'COMPLETED', 'IN_PROGRESS', 'NEW', 'COMPLETED'];

    for (let i = 0; i < extraClients.length; i += 1) {
      const sIdx = i % servicesData.length;
      const st = extraStatuses[i % extraStatuses.length];
      const assigned = st === 'NEW' && i % 2 === 0 ? null : employeeIds[(i % 2) + 1] ? (i % 2) + 1 : 1;
      demoInquiries.push({
        code: `DEMO${String(7 + i).padStart(2, '0')}`,
        serviceIdx: sIdx,
        assignedIdx: assigned,
        contact: extraClients[i],
        subject: `Запит на послугу «${servicesData[sIdx].title}» (об’єкт #${i + 7})`,
        message: `Прошу надати консультацію та попередній розрахунок вартості виконання робіт для об’єкта #${i + 7}.`,
        status: st,
        daysOffsetCreated: -(i % 4),
        dueDateOverride: null,
        publicResp: st === 'COMPLETED' ? 'Роботи виконано у повному обсязі, акт підписано.' : null,
        internalNote: `Внутрішній запис менеджера щодо об’єкта #${i + 7}.`
      });
    }

    const insertContact = db.prepare(
      `INSERT INTO client_contacts (contact_name, company_name, phone, email, created_at)
       VALUES (?, ?, ?, ?, ?)`
    );
    const insertInquiry = db.prepare(
      `INSERT INTO inquiries (
        ticket_number, access_code, service_id, contact_id, assigned_employee_id,
        subject, message, status, rejection_reason, public_response, idempotency_key,
        sla_business_days, due_date, created_at, updated_at, closed_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
    );
    const insertHistory = db.prepare(
      `INSERT INTO inquiry_history (
        inquiry_id, author_employee_id, event_type, old_status, new_status,
        comment, is_internal, created_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
    );

    const year = todayStr.slice(0, 4);
    for (let i = 0; i < demoInquiries.length; i += 1) {
      const d = demoInquiries[i];
      const service = serviceIds[d.serviceIdx];
      const createdAt = shiftDays(todayStr, d.daysOffsetCreated);
      const dueDate =
        d.dueDateOverride ||
        addBusinessDays(createdAt.slice(0, 10), service.sla_business_days, exceptionsMap);

      const contactRes = insertContact.run(
        d.contact[0],
        d.contact[1],
        d.contact[2],
        d.contact[3],
        createdAt
      );
      const contactId = Number(contactRes.lastInsertRowid);
      const assignedId = d.assignedIdx !== null ? employeeIds[d.assignedIdx] : null;
      const closedAt =
        d.status === 'COMPLETED' || d.status === 'REJECTED' ? nowIso : null;
      const ticketNumber = `VP-${year}-${String(i + 1).padStart(4, '0')}`;

      const inqRes = insertInquiry.run(
        ticketNumber,
        d.code,
        service.id,
        contactId,
        assignedId,
        d.subject,
        d.message,
        d.status,
        d.rejectionReason || null,
        d.publicResp || null,
        `seed-${ticketNumber}`,
        service.sla_business_days,
        dueDate,
        createdAt,
        nowIso,
        closedAt
      );
      const inquiryId = Number(inqRes.lastInsertRowid);

      insertHistory.run(
        inquiryId,
        null,
        'CREATED',
        null,
        'NEW',
        `Звернення зареєстровано. Очікуваний строк: ${service.sla_business_days} роб. дн. (до ${dueDate}).`,
        0,
        createdAt
      );

      if (d.status !== 'NEW') {
        insertHistory.run(
          inquiryId,
          assignedId || employeeIds[0],
          'STATUS_CHANGE',
          'NEW',
          d.status,
          d.rejectionReason
            ? `Статус змінено на «Відхилена». Причина: ${d.rejectionReason}`
            : `Статус змінено на «${d.status === 'COMPLETED' ? 'Завершена' : 'В роботі'}».`,
          0,
          nowIso
        );
      }

      if (d.internalNote) {
        insertHistory.run(
          inquiryId,
          assignedId || employeeIds[0],
          'INTERNAL_NOTE',
          null,
          null,
          d.internalNote,
          1,
          nowIso
        );
      }
    }

    return { seeded: true, count: demoInquiries.length };
  });
}

if (require.main === module) {
  const config = loadConfig();
  const db = openDatabase(config.dbPath);
  try {
    const result = seedDatabase(db);
    console.log(result.seeded ? `Створено ${result.count} демонстраційних звернень.` : result.reason);
  } finally {
    db.close();
  }
}

module.exports = { seedDatabase };
