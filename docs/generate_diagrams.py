import json
from pathlib import Path
import matplotlib.pyplot as plt
import matplotlib.patches as patches

plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['axes.unicode_minus'] = False

ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / 'docs' / 'diagrams'
OUT_DIR.mkdir(parents=True, exist_ok=True)


def draw_box(ax, x, y, w, h, text, fc='#eef6f8', ec='#0f4c5c', lw=1.5, fontsize=9.5, bold=False, rounded=False):
    if rounded:
        box = patches.FancyBboxPatch(
            (x, y), w, h,
            boxstyle='round,pad=0.02,rounding_size=0.04',
            facecolor=fc, edgecolor=ec, linewidth=lw
        )
    else:
        box = patches.Rectangle((x, y), w, h, facecolor=fc, edgecolor=ec, linewidth=lw)
    ax.add_patch(box)
    ax.text(
        x + w / 2, y + h / 2, text,
        ha='center', va='center', fontsize=fontsize,
        fontweight='bold' if bold else 'normal', color='#111827'
    )


def draw_diamond(ax, cx, cy, w, h, text, fc='#fff8e6', ec='#b45309', fontsize=9):
    pts = [(cx, cy + h / 2), (cx + w / 2, cy), (cx, cy - h / 2), (cx - w / 2, cy)]
    poly = patches.Polygon(pts, closed=True, facecolor=fc, edgecolor=ec, linewidth=1.5)
    ax.add_patch(poly)
    ax.text(cx, cy, text, ha='center', va='center', fontsize=fontsize, color='#111827')


def draw_arrow(ax, x1, y1, x2, y2, label='', lx=None, ly=None, dashed=False):
    style = '--' if dashed else '-'
    ax.annotate(
        '', xy=(x2, y2), xytext=(x1, y1),
        arrowprops=dict(arrowstyle='->', color='#374151', lw=1.3, linestyle=style)
    )
    if label:
        tx = lx if lx is not None else (x1 + x2) / 2
        ty = ly if ly is not None else (y1 + y2) / 2
        ax.text(
            tx, ty, label, ha='center', va='center', fontsize=8.2,
            color='#1f2937', bbox=dict(boxstyle='round,pad=0.15', fc='white', ec='none', alpha=0.85)
        )


def new_canvas(w=10, h=6.5, xlim=(0, 10), ylim=(0, 6.5)):
    fig, ax = plt.subplots(figsize=(w, h), dpi=220)
    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)
    ax.axis('off')
    return fig, ax


def gen_use_cases():
    fig, ax = new_canvas(10, 6.8, (0, 10), (0, 6.8))
    # Межа системи
    rect = patches.Rectangle((2.5, 0.4), 5.0, 6.0, facecolor='#f9fafb', edgecolor='#0f4c5c', linewidth=1.8)
    ax.add_patch(rect)
    ax.text(5.0, 6.15, 'Вебсистема «ВольтаПро Сервіс»', ha='center', va='center', fontsize=11, fontweight='bold', color='#0f4c5c')

    cases = [
        (5.0, 5.4, 'Перегляд каталогу послуг\nта пошук за фільтрами'),
        (5.0, 4.5, 'Подання звернення\nта отримання коду доступу'),
        (5.0, 3.6, 'Перевірка статусу заявки\nза номером і кодом'),
        (5.0, 2.7, 'Робота зі списком заявок,\nпризначення та коментарі'),
        (5.0, 1.8, 'Зміна статусу та перегляд\nзвіту прострочень SLA'),
        (5.0, 0.9, 'Повернення завершеної заявки\nта перерозподіл між відділами'),
    ]
    for cx, cy, label in cases:
        el = patches.Ellipse((cx, cy), 3.9, 0.68, facecolor='#eef6f8', edgecolor='#0f4c5c', lw=1.3)
        ax.add_patch(el)
        ax.text(cx, cy, label, ha='center', va='center', fontsize=8.8)

    # Актори
    def actor(x, y, name):
        ax.add_patch(patches.Circle((x, y + 0.35), 0.16, fc='white', ec='#111827', lw=1.4))
        ax.plot([x, x], [y + 0.19, y - 0.15], color='#111827', lw=1.4)
        ax.plot([x - 0.22, x + 0.22], [y + 0.08, y + 0.08], color='#111827', lw=1.4)
        ax.plot([x, x - 0.18], [y - 0.15, y - 0.42], color='#111827', lw=1.4)
        ax.plot([x, x + 0.18], [y - 0.15, y - 0.42], color='#111827', lw=1.4)
        ax.text(x, y - 0.62, name, ha='center', va='center', fontsize=9.5, fontweight='bold')

    actor(1.1, 4.5, 'Гість / Клієнт')
    actor(8.9, 2.7, 'Менеджер')
    actor(8.9, 1.0, 'Старший\nменеджер')

    for cy in [5.4, 4.5, 3.6]:
        ax.plot([1.45, 3.05], [4.5, cy], color='#374151', lw=1.2)
    for cy in [5.4, 2.7, 1.8]:
        ax.plot([8.55, 6.95], [2.7, cy], color='#374151', lw=1.2)
    for cy in [2.7, 1.8, 0.9]:
        ax.plot([8.55, 6.95], [1.0, cy], color='#374151', lw=1.2)

    fig.tight_layout()
    fig.savefig(OUT_DIR / 'fig_1_1_use_cases.png', bbox_inches='tight')
    plt.close(fig)


def gen_sla_algorithm():
    fig, ax = new_canvas(9.5, 7.2, (0, 9.5), (0, 7.2))
    draw_box(ax, 3.2, 6.4, 3.0, 0.5, 'Початок\n(startDate, slaDays, exceptions)', fc='#e0f2fe', rounded=True, bold=True)
    draw_box(ax, 3.0, 5.45, 3.4, 0.55, 'cursor = toIsoDate(startDate)\nremaining = slaDays')
    draw_diamond(ax, 4.7, 4.45, 2.6, 0.85, 'remaining > 0 ?')
    draw_box(ax, 3.1, 3.25, 3.2, 0.5, 'cursor = cursor + 1 календарний день')
    draw_diamond(ax, 4.7, 2.2, 3.0, 0.9, 'isWorkingDay(cursor,\nexceptions) ?')
    draw_box(ax, 3.3, 0.95, 2.8, 0.5, 'remaining = remaining - 1')
    draw_box(ax, 6.6, 4.2, 2.5, 0.55, 'Повернути cursor\n(дата дедлайну)', fc='#dcfce7', ec='#15803d', rounded=True, bold=True)
    draw_box(ax, 6.6, 3.1, 2.5, 0.5, 'Кінець', fc='#e0f2fe', rounded=True, bold=True)

    draw_arrow(ax, 4.7, 6.4, 4.7, 6.0)
    draw_arrow(ax, 4.7, 5.45, 4.7, 4.88)
    draw_arrow(ax, 4.7, 4.02, 4.7, 3.75, 'Так', 4.95, 3.9)
    draw_arrow(ax, 4.7, 3.25, 4.7, 2.65)
    draw_arrow(ax, 4.7, 1.75, 4.7, 1.45, 'Так', 4.95, 1.6)
    draw_arrow(ax, 6.0, 4.45, 6.6, 4.45, 'Ні', 6.3, 4.65)
    draw_arrow(ax, 7.85, 4.2, 7.85, 3.6)

    # Зворотні зв'язки зліва
    ax.plot([3.2, 1.6, 1.6, 3.4], [2.2, 2.2, 4.45, 4.45], color='#374151', lw=1.3)
    ax.text(2.4, 2.38, 'Ні (вихідний/свято)', fontsize=8, ha='center')
    ax.plot([3.3, 1.2, 1.2, 3.4], [1.2, 1.2, 4.45, 4.45], color='#374151', lw=1.3)
    ax.annotate('', xy=(3.4, 4.45), xytext=(3.1, 4.45), arrowprops=dict(arrowstyle='->', color='#374151', lw=1.3))

    fig.tight_layout()
    fig.savefig(OUT_DIR / 'fig_2_1_sla_algorithm.png', bbox_inches='tight')
    plt.close(fig)


def gen_status_algorithm():
    fig, ax = new_canvas(9.5, 6.8, (0, 9.5), (0, 6.8))
    draw_box(ax, 3.0, 6.0, 3.5, 0.5, 'Початок\n(current, next, role, reason)', fc='#e0f2fe', rounded=True, bold=True)
    draw_diamond(ax, 4.75, 5.0, 3.2, 0.85, 'Перехід у таблиці\nALLOWED_TRANSITIONS?')
    draw_diamond(ax, 4.75, 3.75, 3.4, 0.85, 'current IN {COMPLETED, REJECTED}\nТА role != SENIOR_MANAGER?')
    draw_diamond(ax, 4.75, 2.5, 3.2, 0.85, 'next == REJECTED ТА\nlength(reason.trim()) < 5?')
    draw_box(ax, 3.2, 1.15, 3.1, 0.55, 'Оновити статус, closed_at\nі записати в inquiry_history', fc='#dcfce7', ec='#15803d')
    draw_box(ax, 3.75, 0.3, 2.0, 0.45, 'Кінець (200 OK)', fc='#e0f2fe', rounded=True, bold=True)

    draw_box(ax, 7.1, 4.75, 2.1, 0.5, 'Відмова 409\nConflict', fc='#fee2e2', ec='#b91c1c')
    draw_box(ax, 7.1, 3.5, 2.1, 0.5, 'Відмова 403\nForbidden', fc='#fee2e2', ec='#b91c1c')
    draw_box(ax, 7.1, 2.25, 2.1, 0.5, 'Відмова 400\nBad Request', fc='#fee2e2', ec='#b91c1c')

    draw_arrow(ax, 4.75, 6.0, 4.75, 5.42)
    draw_arrow(ax, 4.75, 4.58, 4.75, 4.18, 'Так', 5.0, 4.38)
    draw_arrow(ax, 4.75, 3.32, 4.75, 2.92, 'Ні', 5.0, 3.12)
    draw_arrow(ax, 4.75, 2.08, 4.75, 1.7, 'Ні', 5.0, 1.9)
    draw_arrow(ax, 4.75, 1.15, 4.75, 0.75)

    draw_arrow(ax, 6.35, 5.0, 7.1, 5.0, 'Ні', 6.72, 5.18)
    draw_arrow(ax, 6.45, 3.75, 7.1, 3.75, 'Так', 6.77, 3.93)
    draw_arrow(ax, 6.35, 2.5, 7.1, 2.5, 'Так', 6.72, 2.68)

    fig.tight_layout()
    fig.savefig(OUT_DIR / 'fig_2_2_status_transition.png', bbox_inches='tight')
    plt.close(fig)


def gen_ui_state_map():
    fig, ax = new_canvas(10.5, 6.2, (0, 10.5), (0, 6.2))
    draw_box(ax, 0.4, 4.6, 2.7, 0.9, 'UI-01: Публічний сайт\nКаталог послуг + фільтри\n(порожній / список / помилка)', bold=True)
    draw_box(ax, 4.0, 4.6, 2.6, 0.9, 'UI-02: Форма звернення\nВведення даних\n(збереження при 400)', bold=True)
    draw_box(ax, 7.5, 4.6, 2.6, 0.9, 'UI-03: Підтвердження\nНомер VP-..., код доступу\nта дедлайн SLA', fc='#dcfce7', ec='#15803d', bold=True)

    draw_box(ax, 4.0, 2.5, 2.6, 0.9, 'UI-04: Перевірка статусу\nПошук за номером + кодом\n(публічна історія)', bold=True)
    draw_box(ax, 0.4, 2.5, 2.7, 0.9, 'UI-05: Вхід менеджера\nКорпоративна пошта і пароль\n(HttpOnly cookie)', bold=True)
    draw_box(ax, 0.4, 0.5, 2.7, 0.95, 'UI-06: Робочий список\nФільтри, пагінація,\nіндикація SLA', bold=True)
    draw_box(ax, 4.0, 0.5, 2.6, 0.95, 'UI-07: Картка звернення\nСтатус, призначення,\nвнутрішні примітки', bold=True)
    draw_box(ax, 7.5, 0.5, 2.6, 0.95, 'UI-08: Звіт SLA\nПрострочені заявки та\nвиробничий календар', fc='#fff8e6', ec='#b45309', bold=True)

    draw_arrow(ax, 3.1, 5.05, 4.0, 5.05, 'обрати послугу')
    draw_arrow(ax, 6.6, 5.05, 7.5, 5.05, '201 Created')
    draw_arrow(ax, 8.8, 4.6, 5.3, 3.4, 'перейти до статусу')
    draw_arrow(ax, 1.75, 4.6, 1.75, 3.4, 'кабінет')
    draw_arrow(ax, 1.75, 2.5, 1.75, 1.45, 'успішний вхід')
    draw_arrow(ax, 3.1, 0.97, 4.0, 0.97, 'відкрити')
    draw_arrow(ax, 3.1, 1.25, 7.5, 1.25, 'вкладка SLA')

    fig.tight_layout()
    fig.savefig(OUT_DIR / 'fig_3_1_ui_state_map.png', bbox_inches='tight')
    plt.close(fig)


def gen_logical_view():
    fig, ax = new_canvas(10, 6.0, (0, 10), (0, 6.0))
    draw_box(ax, 0.5, 4.8, 4.0, 0.8, 'Клієнт (Гість / Заявник)\nПерегляд послуг, подання, перевірка статусу')
    draw_box(ax, 5.5, 4.8, 4.0, 0.8, 'Персонал (Менеджер / Старший менеджер)\nОброблення звернень, примітки, звіт SLA')

    draw_box(ax, 1.5, 3.3, 7.0, 0.8, 'Браузер (HTML5 + CSS3 + JS + Bootstrap 5)\nКлієнтська валідація, стан UI, fetch до REST API', fc='#e0f2fe', bold=True)

    # Довірена серверна межа
    trust = patches.Rectangle((1.0, 1.35), 8.0, 1.35, facecolor='#f8fafc', edgecolor='#0f4c5c', lw=1.6, ls='--')
    ax.add_patch(trust)
    ax.text(5.0, 2.5, 'Довірена серверна частина (Node.js + Express)', ha='center', va='center', fontsize=9.5, fontweight='bold', color='#0f4c5c')
    draw_box(ax, 1.3, 1.55, 2.3, 0.72, 'Auth & RBAC\nscrypt, cookie-сесії')
    draw_box(ax, 3.85, 1.55, 2.3, 0.72, 'Сервіси звернень\nІдемпотентність, переходи')
    draw_box(ax, 6.4, 1.55, 2.3, 0.72, 'Двигун SLA\nКалендар робочих днів')

    draw_box(ax, 2.2, 0.25, 5.6, 0.75, 'Приватне сховище даних (SQLite у WAL-режимі)\nТаблиці, зовнішні ключі, обмеження CHECK, індекси, транзакції', fc='#fef3c7', ec='#b45309', bold=True)

    draw_arrow(ax, 2.5, 4.8, 3.5, 4.1)
    draw_arrow(ax, 7.5, 4.8, 6.5, 4.1)
    draw_arrow(ax, 5.0, 3.3, 5.0, 2.7, 'HTTP / JSON (REST API)')
    draw_arrow(ax, 5.0, 1.35, 5.0, 1.0, 'SQL / Транзакції')

    fig.tight_layout()
    fig.savefig(OUT_DIR / 'fig_3_2_logical_view.png', bbox_inches='tight')
    plt.close(fig)


def gen_deployment_view():
    fig, ax = new_canvas(9.5, 5.8, (0, 9.5), (0, 5.8))
    # Вузол 1
    ax.add_patch(patches.Rectangle((0.6, 3.6), 8.3, 1.7, fc='#f8fafc', ec='#1f2937', lw=1.5))
    ax.text(4.75, 5.05, '<<device>> Пристрій користувача (Windows 11 Pro)', ha='center', va='center', fontsize=9.5, fontweight='bold')
    draw_box(ax, 2.0, 3.85, 5.5, 0.85, '<<executionEnvironment>> Веббраузер Google Chrome\nHTML5 / CSS3 / ES6+ модулі / Bootstrap 5')

    # Вузол 2
    ax.add_patch(patches.Rectangle((0.6, 0.4), 8.3, 2.5, fc='#f8fafc', ec='#0f4c5c', lw=1.6))
    ax.text(4.75, 2.65, '<<node>> Навчальний сервер / Локальний стенд або Docker-контейнер', ha='center', va='center', fontsize=9.5, fontweight='bold', color='#0f4c5c')
    draw_box(ax, 1.0, 0.75, 3.5, 1.5, '<<executionEnvironment>> Node.js v26.13.1\nExpress 4.21 (порт 3000)\nREST API + статика + сесії')
    draw_box(ax, 5.2, 0.75, 3.3, 1.5, '<<executionEnvironment>> SQLite (node:sqlite)\nФайл: data/voltapro.db\nWAL-режим + міграції + seed', fc='#fef3c7', ec='#b45309')

    draw_arrow(ax, 4.75, 3.6, 2.75, 2.25, 'HTTP (localhost:3000)', 3.6, 3.15)
    draw_arrow(ax, 4.5, 1.5, 5.2, 1.5, 'Файловий\nвиклик')

    fig.tight_layout()
    fig.savefig(OUT_DIR / 'fig_3_3_deployment_view.png', bbox_inches='tight')
    plt.close(fig)


def gen_design_layers():
    fig, ax = new_canvas(9.5, 6.0, (0, 9.5), (0, 6.0))
    layers = [
        (4.6, '#e0f2fe', '#0369a1', '1. Шар подання (Routes & Middleware)', 'publicRoutes.js | authRoutes.js | managerRoutes.js | auth.js | errorHandler.js'),
        (3.25, '#eef6f8', '#0f4c5c', '2. Прикладний шар (Application Services)', 'inquiryService.js | authService.js (координація сценаріїв, формування DTO)'),
        (1.9, '#dcfce7', '#15803d', '3. Шар предметних правил (Domain Layer)', 'inquiryRules.js | calendarRules.js | validationRules.js (інваріанти й SLA)'),
        (0.55, '#fef3c7', '#b45309', '4. Інфраструктурні адаптери (Repositories & DB)', 'inquiryRepository.js | serviceRepository.js | employeeRepository.js | SQLite / logger.js')
    ]
    for y, fc, ec, title, desc in layers:
        draw_box(ax, 0.8, y, 7.9, 0.95, f'{title}\n{desc}', fc=fc, ec=ec, fontsize=9.2, bold=False)

    draw_arrow(ax, 4.75, 4.6, 4.75, 4.2, 'виклик сценарію')
    draw_arrow(ax, 4.75, 3.25, 4.75, 2.85, 'перевірка правил')
    draw_arrow(ax, 4.75, 1.9, 4.75, 1.5, 'збереження / читання')

    fig.tight_layout()
    fig.savefig(OUT_DIR / 'fig_3_4_design_layers.png', bbox_inches='tight')
    plt.close(fig)


def gen_sequence_create():
    fig, ax = new_canvas(10, 6.6, (0, 10), (0, 6.6))
    actors = [(1.2, 'Браузер'), (3.6, 'API (Routes)'), (6.2, 'InquiryService'), (8.7, 'SQLite DB')]
    for x, name in actors:
        draw_box(ax, x - 0.85, 5.85, 1.7, 0.48, name, fc='#eef6f8', bold=True)
        ax.plot([x, x], [0.4, 5.85], color='#6b7280', lw=1.1, ls='--')

    draw_arrow(ax, 1.2, 5.35, 3.6, 5.35, 'POST /api/inquiries', 2.4, 5.53)
    draw_arrow(ax, 3.6, 4.85, 6.2, 4.85, 'createPublicInquiry(body)', 4.9, 5.03)
    draw_arrow(ax, 6.2, 4.35, 8.7, 4.35, 'findByIdempotencyKey(key)', 7.45, 4.53)

    # Блок alt
    ax.add_patch(patches.Rectangle((0.4, 0.6), 9.1, 3.35, fc='none', ec='#374151', lw=1.3))
    ax.text(0.7, 3.78, 'alt', fontsize=9, fontweight='bold', bbox=dict(fc='#e5e7eb', ec='#374151', pad=2))
    ax.text(2.4, 3.78, '[повторний запит із тим самим ключем]', fontsize=8.5)
    draw_arrow(ax, 8.7, 3.45, 6.2, 3.45, 'наявний запис', 7.45, 3.6, dashed=True)
    draw_arrow(ax, 3.6, 3.15, 1.2, 3.15, '200 OK (duplicate: true)', 2.4, 3.3, dashed=True)

    ax.plot([0.4, 9.5], [2.85, 2.85], color='#374151', lw=1.0, ls='--')
    ax.text(2.4, 2.65, '[новий коректний запит]', fontsize=8.5)
    draw_arrow(ax, 6.2, 2.3, 8.7, 2.3, 'BEGIN IMMEDIATE; INSERT contact, inquiry, history; COMMIT', 7.45, 2.46)
    draw_arrow(ax, 8.7, 1.8, 6.2, 1.8, 'створений запис VP-2026-...', 7.45, 1.96, dashed=True)
    draw_arrow(ax, 3.6, 1.45, 1.2, 1.45, '201 Created (ticketNumber, accessCode, dueDate)', 2.4, 1.61, dashed=True)

    ax.plot([0.4, 9.5], [1.15, 1.15], color='#374151', lw=1.0, ls='--')
    ax.text(2.2, 0.98, '[помилка валідації полів]', fontsize=8.5)
    draw_arrow(ax, 3.6, 0.72, 1.2, 0.72, '400 Bad Request (RFC 9457 fieldErrors)', 2.4, 0.87, dashed=True)

    fig.tight_layout()
    fig.savefig(OUT_DIR / 'fig_3_5_sequence_create.png', bbox_inches='tight')
    plt.close(fig)


def gen_er_model():
    fig, ax = new_canvas(10.5, 6.8, (0, 10.5), (0, 6.8))
    draw_box(ax, 0.4, 4.9, 2.8, 1.4, 'services\nPK id\nslug, title, category\nbase_price_uah, sla_business_days')
    draw_box(ax, 0.4, 2.5, 2.8, 1.4, 'client_contacts\nPK id\ncontact_name, company_name\nphone, email, created_at')
    draw_box(ax, 4.0, 2.5, 3.2, 2.4, 'inquiries\nPK id | ticket_number, access_code\nFK service_id, contact_id\nFK assigned_employee_id\nsubject, status, due_date\nrejection_reason, idempotency_key', fc='#e0f2fe')
    draw_box(ax, 4.0, 0.35, 3.2, 1.4, 'inquiry_history\nPK id | FK inquiry_id\nFK author_employee_id\nevent_type, comment, is_internal')
    draw_box(ax, 7.9, 4.7, 2.3, 1.5, 'employees\nPK id\nfull_name, email\npassword_hash, role')
    draw_box(ax, 7.9, 2.6, 2.3, 1.2, 'sessions\nPK token\nFK employee_id\nexpires_at')
    draw_box(ax, 7.9, 0.45, 2.3, 1.2, 'calendar_exceptions\nPK date\nis_working_day\ndescription', fc='#fff8e6', ec='#b45309')

    ax.plot([3.2, 4.0], [5.4, 4.3], color='#374151', lw=1.3)
    ax.text(3.45, 5.0, '1 : 0..*', fontsize=8.5)
    ax.plot([3.2, 4.0], [3.2, 3.2], color='#374151', lw=1.3)
    ax.text(3.6, 3.35, '1 : 0..*', fontsize=8.5)
    ax.plot([5.6, 5.6], [2.5, 1.75], color='#374151', lw=1.3)
    ax.text(5.95, 2.1, '1 : 1..*', fontsize=8.5)
    ax.plot([7.9, 7.2], [5.2, 4.2], color='#374151', lw=1.3)
    ax.text(7.55, 4.9, '1 : 0..*', fontsize=8.5)
    ax.plot([9.05, 9.05], [4.7, 3.8], color='#374151', lw=1.3)
    ax.text(9.35, 4.25, '1 : 0..*', fontsize=8.5)

    fig.tight_layout()
    fig.savefig(OUT_DIR / 'fig_3_6_er_model.png', bbox_inches='tight')
    plt.close(fig)


def gen_domain_dto():
    fig, ax = new_canvas(10.5, 6.2, (0, 10.5), (0, 6.2))
    ax.text(2.6, 5.9, 'Доменна модель та правила', ha='center', fontsize=10.5, fontweight='bold', color='#0f4c5c')
    ax.text(7.8, 5.9, 'Структури передавання даних (DTO)', ha='center', fontsize=10.5, fontweight='bold', color='#b45309')
    ax.plot([5.25, 5.25], [0.3, 6.0], color='#9ca3af', lw=1.2, ls='--')

    draw_box(ax, 0.4, 3.8, 2.1, 1.6, 'Service\nid, title, category\nslaBusinessDays\nisActive')
    draw_box(ax, 2.8, 3.8, 2.1, 1.6, 'Employee\nid, fullName, role\nisActive')
    draw_box(ax, 1.0, 1.2, 3.2, 1.9, 'Inquiry\nid, ticketNumber, status\ndueDate, rejectionReason\nvalidateTransition()\nevaluateSlaState()', fc='#e0f2fe')

    ax.plot([1.45, 2.0], [3.8, 3.1], color='#374151', lw=1.2)
    ax.plot([3.85, 3.2], [3.8, 3.1], color='#374151', lw=1.2)

    draw_box(ax, 5.7, 4.0, 2.1, 1.5, '<<DTO>>\nCreateInquiryRequest\nserviceId, contactName\nphone, email, subject\nmessage, idempotencyKey', fc='#fff8e6', ec='#b45309')
    draw_box(ax, 8.1, 4.0, 2.1, 1.5, '<<DTO>>\nPublicInquiryDto\nticketNumber, status\ndueDate, publicResponse\npublicHistory[]', fc='#fff8e6', ec='#b45309')
    draw_box(ax, 5.7, 1.3, 2.1, 1.8, '<<DTO>>\nManagerDetailDto\nid, ticketNumber, contact\nstatus, slaState\nhistory[] (з internal)', fc='#fff8e6', ec='#b45309')
    draw_box(ax, 8.1, 1.3, 2.1, 1.8, '<<DTO>>\nProblemDetailsError\ntype, title, status\ncode, fieldErrors[]\nrequestId', fc='#fee2e2', ec='#b91c1c')

    fig.tight_layout()
    fig.savefig(OUT_DIR / 'fig_3_7_domain_dto.png', bbox_inches='tight')
    plt.close(fig)


def gen_architecture_classes():
    fig, ax = new_canvas(9.0, 6.2, (0, 9.0), (0, 6.2))
    draw_box(ax, 2.4, 5.0, 4.2, 0.9, 'ManagerRoutes / PublicRoutes\nОбробка HTTP, валідація query/body, статус-коди', bold=True)
    draw_box(ax, 2.4, 3.55, 4.2, 0.95, 'InquiryService\ncreatePublicInquiry(), lookupPublicInquiry(),\nupdateInquiryStatus(), getOverdueReport()', fc='#e0f2fe')
    draw_box(ax, 0.4, 1.9, 3.6, 1.0, 'DomainRules\nvalidateStatusTransition(),\naddBusinessDays(), evaluateSlaState()', fc='#dcfce7', ec='#15803d')
    draw_box(ax, 4.9, 1.9, 3.7, 1.0, '<<interface / contract>>\nInquiryRepository\ncreateWithContactAndHistory(), listForManager()')
    draw_box(ax, 4.9, 0.45, 3.7, 0.9, 'SQLiteInquiryRepository\nnode:sqlite DatabaseSync + BEGIN IMMEDIATE', fc='#fef3c7', ec='#b45309')

    draw_arrow(ax, 4.5, 5.0, 4.5, 4.5, 'викликає', dashed=True)
    draw_arrow(ax, 3.5, 3.55, 2.2, 2.9, 'перевіряє правила', dashed=True)
    draw_arrow(ax, 5.5, 3.55, 6.75, 2.9, 'залежить від', dashed=True)
    draw_arrow(ax, 6.75, 1.35, 6.75, 1.9, 'реалізація', dashed=True)

    fig.tight_layout()
    fig.savefig(OUT_DIR / 'fig_4_1_architecture_classes.png', bbox_inches='tight')
    plt.close(fig)


def gen_perf_charts():
    load_path = ROOT / 'evidence' / 'reports' / 'load-report.json'
    vol_path = ROOT / 'evidence' / 'reports' / 'volume-report.json'

    if load_path.exists():
        load_data = json.loads(load_path.read_text(encoding='utf-8'))['stages']
        vus = [s['virtualUsers'] for s in load_data]
        p50 = [s['p50Ms'] for s in load_data]
        p95 = [s['p95Ms'] for s in load_data]
        p99 = [s['p99Ms'] for s in load_data]
        rps = [s['rps'] for s in load_data]

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10.2, 4.2), dpi=220)
        ax1.plot(vus, p50, marker='o', label='p50 (мс)', lw=2)
        ax1.plot(vus, p95, marker='s', label='p95 (мс)', lw=2)
        ax1.plot(vus, p99, marker='^', label='p99 (мс)', lw=2)
        ax1.set_title('Затримка відповіді API (мс)')
        ax1.set_xlabel('Віртуальні користувачі (VUs)')
        ax1.set_ylabel('Час відповіді, мс')
        ax1.set_xticks(vus)
        ax1.grid(True, linestyle='--', alpha=0.5)
        ax1.legend()

        ax2.bar([str(v) for v in vus], rps, color='#0f4c5c', width=0.5)
        ax2.set_title('Пропускна здатність (запитів/с)')
        ax2.set_xlabel('Віртуальні користувачі (VUs)')
        ax2.set_ylabel('RPS (req/s)')
        ax2.grid(True, axis='y', linestyle='--', alpha=0.5)
        for idx, val in enumerate(rps):
            ax2.text(idx, val + 15, f'{val:.1f}', ha='center', fontsize=9)

        fig.tight_layout()
        fig.savefig(OUT_DIR / 'fig_5_1_load_chart.png', bbox_inches='tight')
        plt.close(fig)

    if vol_path.exists():
        vol_data = json.loads(vol_path.read_text(encoding='utf-8'))['results']
        sizes = [str(r['recordsCount']) for r in vol_data]
        list_ms = [r['listFiltered']['avgMs'] for r in vol_data]
        search_ms = [r['searchByText']['avgMs'] for r in vol_data]
        sla_ms = [r['overdueReport']['avgMs'] for r in vol_data]

        fig, ax = plt.subplots(figsize=(8.5, 4.2), dpi=220)
        x = range(len(sizes))
        w = 0.25
        ax.bar([i - w for i in x], list_ms, width=w, label='Фільтрований список', color='#0f4c5c')
        ax.bar(x, search_ms, width=w, label='Текстовий пошук', color='#0284c7')
        ax.bar([i + w for i in x], sla_ms, width=w, label='Звіт прострочень SLA', color='#e36414')
        ax.set_xticks(list(x))
        ax.set_xticklabels([f'{s} записів' for s in sizes])
        ax.set_ylabel('Середній час, мс')
        ax.set_title('Вплив обсягу даних у SQLite на час виконання запитів')
        ax.grid(True, axis='y', linestyle='--', alpha=0.5)
        ax.legend()

        fig.tight_layout()
        fig.savefig(OUT_DIR / 'fig_5_2_volume_chart.png', bbox_inches='tight')
        plt.close(fig)


if __name__ == '__main__':
    gen_use_cases()
    gen_sla_algorithm()
    gen_status_algorithm()
    gen_ui_state_map()
    gen_logical_view()
    gen_deployment_view()
    gen_design_layers()
    gen_sequence_create()
    gen_er_model()
    gen_domain_dto()
    gen_architecture_classes()
    gen_perf_charts()
    print('Усі 13 діаграм і графіків згенеровано у', OUT_DIR)
