import json
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

ROOT = Path(__file__).resolve().parent.parent
DIAG_DIR = ROOT / 'docs' / 'diagrams'
SCR_DIR = ROOT / 'evidence' / 'screenshots'
REP_DIR = ROOT / 'evidence' / 'reports'
OUT_DOCX = ROOT / 'docs' / 'Курсова_робота_Жиган_Д_О_АІ-245.docx'


def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="4B5563"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="4B5563"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="4B5563"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="4B5563"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="9CA3AF"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="9CA3AF"/>'
        '</w:tblBorders>'
    )
    tblPr.append(borders)


def add_page_number_to_footer(section):
    section.different_first_page_header_footer = True
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run()
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    run._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w')))
    run._r.append(parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w')))
    run._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w')))
    run._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w')))


def configure_styles(doc: Document):
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(1.5)
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(2.0)
    add_page_number_to_footer(sec)

    normal = doc.styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(14)
    normal.font.color.rgb = RGBColor(0, 0, 0)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.first_line_indent = Cm(1.25)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    for lvl in [1, 2, 3]:
        st = doc.styles[f'Heading {lvl}']
        st.font.name = 'Times New Roman'
        st.font.size = Pt(14)
        st.font.bold = True
        st.font.color.rgb = RGBColor(0, 0, 0)
        st.paragraph_format.line_spacing = 1.5
        st.paragraph_format.keep_with_next = True
        st.paragraph_format.space_before = Pt(10 if lvl > 1 else 0)
        st.paragraph_format.space_after = Pt(6)
        if lvl == 1:
            st.paragraph_format.first_line_indent = Cm(0)
            st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            st.paragraph_format.first_line_indent = Cm(1.25)
            st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT


def add_p(doc, text, bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=1.25, space_after=0):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.first_line_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)
    r.bold = bold
    return p


def add_bullet(doc, text, bold_prefix=''):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Cm(1.25)
    p.paragraph_format.line_spacing = 1.5
    r0 = p.add_run('– ')
    r0.font.name = 'Times New Roman'
    r0.font.size = Pt(14)
    if bold_prefix:
        rb = p.add_run(bold_prefix)
        rb.font.name = 'Times New Roman'
        rb.font.size = Pt(14)
        rb.bold = True
    r1 = p.add_run(text)
    r1.font.name = 'Times New Roman'
    r1.font.size = Pt(14)
    return p


def add_h1(doc, text, page_break=True):
    if page_break:
        doc.add_page_break()
    return doc.add_heading(text, level=1)


def add_h2(doc, text):
    return doc.add_heading(text, level=2)


def add_code_block(doc, title, code_text):
    add_p(doc, title, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, indent=1.25, space_after=3)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(code_text.strip())
    r.font.name = 'Courier New'
    r.font.size = Pt(11)


def add_table(doc, caption, headers, rows, col_widths=None):
    if caption:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(caption)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(14)

    tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl)

    # Повтор заголовка на новій сторінці, БЕЗ сірої заливки (чистий білий фон)
    trPr = tbl.rows[0]._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

    for j, h in enumerate(headers):
        cell = tbl.rows[0].cells[j]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.line_spacing = 1.0
        r = p.add_run(str(h))
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12)
        r.bold = True

    for i, row in enumerate(rows):
        trPr_r = tbl.rows[i + 1]._tr.get_or_add_trPr()
        trPr_r.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for j, val in enumerate(row):
            cell = tbl.rows[i + 1].cells[j]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.line_spacing = 1.0
            r = p.add_run(str(val))
            r.font.name = 'Times New Roman'
            r.font.size = Pt(12)

    if col_widths:
        for row in tbl.rows:
            for j, w in enumerate(col_widths):
                row.cells[j].width = Cm(w)

    sp = doc.add_paragraph()
    sp.paragraph_format.first_line_indent = Cm(0)
    sp.paragraph_format.space_after = Pt(4)
    return tbl


def add_figure(doc, img_path: Path, caption: str, width_cm=15.5):
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.first_line_indent = Cm(0)
    p_img.paragraph_format.space_before = Pt(6)
    p_img.paragraph_format.space_after = Pt(3)
    p_img.paragraph_format.keep_with_next = True
    run = p_img.add_run()
    run.add_picture(str(img_path), width=Cm(width_cm))

    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.first_line_indent = Cm(0)
    p_cap.paragraph_format.space_after = Pt(8)
    r = p_cap.add_run(caption)
    r.font.name = 'Times New Roman'
    r.font.size = Pt(12)


def add_toc_field(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run()
    run._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w')))
    run._r.append(parse_xml(r'<w:instrText %s xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText>' % nsdecls('w')))
    run._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w')))
    run._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w')))


def build_document():
    doc = Document()
    configure_styles(doc)

    load_rep = json.loads((REP_DIR / 'load-report.json').read_text(encoding='utf-8'))['stages']
    vol_rep = json.loads((REP_DIR / 'volume-report.json').read_text(encoding='utf-8'))['results']
    code_rep = json.loads((REP_DIR / 'code-metrics.json').read_text(encoding='utf-8'))
    e2e_rep = json.loads((REP_DIR / 'e2e-ui-report.json').read_text(encoding='utf-8'))['checks']

    # =========================================================================
    # ТИТУЛЬНИЙ АРКУШ
    # =========================================================================
    for line in [
        'Міністерство освіти і науки України',
        'Національний університет «Одеська політехніка»',
        'Навчально-науковий інститут комп’ютерних систем',
        'Кафедра інформаційних систем'
    ]:
        add_p(doc, line, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=2)

    add_p(doc, '', indent=0, space_after=44)
    add_p(doc, 'КУРСОВА РОБОТА', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=6)
    add_p(doc, 'з дисципліни «Веб-технології та веб-дизайн»', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=6)
    add_p(doc, 'на тему:', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=6)
    add_p(
        doc,
        '«Корпоративний сайт із керуванням зверненнями для підприємства «ВольтаПро Сервіс» із контролем строків опрацювання SLA та рольовим розмежуванням доступу»',
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        indent=0,
        space_after=8
    )
    add_p(doc, 'Варіант 5 (Група 2. Корпоративні вебсистеми)', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=24)

    add_p(doc, 'Спеціальність: 122 / F3 «Комп’ютерні науки»', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=4)
    add_p(doc, 'Освітня програма: «Комп’ютерні науки»', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=12)
    add_p(doc, 'Виконала: здобувачка групи АІ-245', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=2)
    add_p(doc, 'ПІБ: Жиган Дар’я Олександрівна', bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=4)
    add_p(doc, 'Підпис ________________________     Дата ______________________', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=12)
    add_p(doc, 'Керівник: Червоненко Петро Петрович', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=4)
    add_p(doc, 'Підпис ________________________     Дата ______________________', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=40)
    add_p(doc, 'Одеса — 2026', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=0)

    # =========================================================================
    # ЗАВДАННЯ НА КУРСОВУ РОБОТУ
    # =========================================================================
    doc.add_page_break()
    add_p(doc, 'Національний університет «Одеська політехніка»', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=2)
    add_p(doc, 'Кафедра інформаційних систем', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=6)
    add_p(doc, 'ЗАВДАННЯ НА КУРСОВУ РОБОТУ', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=2)
    add_p(doc, 'з дисципліни «Веб-технології та веб-дизайн»', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=8)

    add_p(doc, 'Здобувачка: Жиган Дар’я Олександрівна     Група: АІ-245     Спеціальність: 122 / F3 «Комп’ютерні науки»', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=2)
    add_p(doc, 'Керівник: Червоненко Петро Петрович', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=6)
    add_p(doc, '1. Тема роботи', bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=2)
    add_p(doc, 'Корпоративний сайт із керуванням зверненнями для підприємства «ВольтаПро Сервіс» із контролем строків опрацювання SLA та рольовим розмежуванням доступу (Варіант 5).', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=4)

    add_table(
        doc,
        '2. Мета, користувачі та вихідні дані',
        ['Параметр', 'Зміст індивідуального завдання'],
        [
            ['Мета розроблення', 'Забезпечити прозорий та керований процес приймання, розподілу й опрацювання клієнтських звернень підприємства «ВольтаПро Сервіс» через вебінтерфейс із регламентованим контролем строків SLA у робочих днях, захищеним пошуком заявки клієнтом та рольовим доступом персоналу.'],
            ['Ролі користувачів', 'Гість / Клієнт (перегляд каталогу послуг, подання звернення, пошук заявки за номером і кодом доступу); Менеджер (MANAGER — робота зі своїми та нерозподіленими заявками, внутрішні примітки, зміна статусу); Старший менеджер (SENIOR_MANAGER — перерозподіл між відділами, повернення завершених/відхилених заявок у роботу, повний звіт SLA).'],
            ['Основні дані та джерела', 'Синтетичний навчальний набір: 8 послуг у 4 категоріях, 3 облікові записи працівників, 4 винятки виробничого календаря, 18 початкових звернень у різних статусах та станах SLA (міграція 001_init.sql та ідемпотентний seed.js).'],
            ['Обов’язковий наскрізний сценарій', 'Пошук і фільтрація послуг → вибір послуги → подання звернення (із захистом від дублювання та збереженням полів при 400) → отримання номера та коду доступу → вхід менеджера → призначення відповідального → додавання внутрішньої примітки й публічної відповіді → зміна статусу → перевірка заявки клієнтом без витоку внутрішніх приміток.'],
            ['Погоджене ускладнення', 'Контроль строку опрацювання (SLA) у робочих днях за явно визначеним виробничим календарем (вихідні субота/неділя + таблиця святкових і перенесених робочих днів calendar_exceptions) та інтерактивний звіт про прострочені заявки.'],
            ['Поза межами роботи', 'Реальне надсилання SMS та електронних листів; інтеграція із зовнішніми CRM/ERP чи платіжними шлюзами; цілодобова промислова експлуатація.']
        ],
        col_widths=[4.8, 12.2]
    )

    add_table(
        doc,
        '3. Обов’язкові функції та критерії приймання',
        ['Код / функція', 'Очікуваний результат і спосіб перевірки'],
        [
            ['FR-01 / Каталог послуг', 'Пошук (включно з кирилицею без урахування регістру через uk_lower) та фільтрація за категорією повертають лише активні послуги; при відсутності збігів — порожній стан. Перевірка: інтеграційний тест + E2E.'],
            ['FR-02, FR-03 / Форма та дублікати', 'Неповний контакт відхиляється (400 VALIDATION_ERROR) зі збереженням полів форми; повторний запит із тим самим idempotencyKey повертає 200 без створення дубліката. Перевірка: unit + integration + E2E.'],
            ['FR-04 / Пошук клієнтом', 'Пошук за ticketNumber + 6-символьним accessCode повертає статус, дедлайн і лише публічну історію (is_internal = 0); невірний код → 404. Перевірка: integration.test.js + E2E.'],
            ['FR-05, FR-06 / Статуси й відмова', 'Переходи NEW → IN_PROGRESS → COMPLETED / REJECTED; відхилення без причини (< 5 симв.) повертає 400 REJECTION_REASON_REQUIRED; повернення завершеної заявки заборонено для MANAGER (403) і дозволено SENIOR_MANAGER.'],
            ['FR-07 / Внутрішні примітки', 'Службові примітки (is_internal = 1) зберігаються в inquiry_history та відображаються лише в кабінеті менеджера, але ніколи не потрапляють у публічний DTO.'],
            ['FR-08 / Контроль SLA та звіт', 'Дедлайн due_date обчислюється за робочими днями календаря; звіт /api/manager/reports/overdue групує активні, прострочені та «дедлайн сьогодні» заявки. Перевірка: unit.test.js + integration.test.js.'],
            ['NFR-SEC-01 / Безпека та доступ', 'Паролі хешуються scrypt із сіллю (21 символ); сесії у HttpOnly cookie; гість без сесії отримує 401; manager2 при доступі до чужої заявки manager1 отримує 403.'],
            ['NFR-UI-01, NFR-PERF-01 / UI й навантаження', 'Відсутність горизонтального переповнення на 360/768/1280 px, навігація з клавіатури; p95 API ≤ 800 мс, помилок < 1 % при навантаженні до 25 VU.']
        ],
        col_widths=[4.8, 12.2]
    )

    add_p(doc, '4. Склад пояснювальної записки, графічні матеріали та календарний план', bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=2)
    add_p(doc, 'Стек та умови перевірки: Node.js v26.13.1, Express 4.21.2, вбудований node:sqlite (WAL-режим), HTML5, CSS3, ES6+ JavaScript, Bootstrap 5 (локально), Docker Compose, вбудований test runner (node:test), Playwright + Google Chrome.', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=4)
    add_p(doc, 'Графічні матеріали (PNG у docs/diagrams/ та скриншоти у evidence/screenshots/): UML-діаграма прецедентів, блок-схеми алгоритмів SLA і переходів статусів, карта станів UI, структурна схема, діаграма розгортання, схема шарів, діаграма послідовності, ER-модель, діаграми доменних класів і DTO, діаграма модулів реалізації та графіки тестування.', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=4)

    add_table(
        doc,
        '',
        ['Етап', 'Тижні', 'Результат виконання'],
        [
            ['1. Аналіз предметної області, аналогів і вимог', '1–4', 'Погоджене завдання (28.10.2026), каталог вимог FR/NFR, діаграма прецедентів'],
            ['2. Алгоритми, архітектура, UI-макети та модель БД', '4–6', 'Блок-схеми SLA, ER-схема 001_init.sql, архітектурні подання'],
            ['3. Реалізація вебсистеми, безпеки та ускладнень', '6–10', 'Серверне API, клієнтський UI на Bootstrap 5, календар SLA і звіт прострочень'],
            ['4. Тестування, оптимізація та оформлення записки', '10–14', 'Автотести (90,00 % покриття), E2E Chrome, навантаження, DOCX і PDF (27.11.2026)']
        ],
        col_widths=[6.2, 2.0, 8.8]
    )

    add_p(doc, 'Репозиторій: GitHub https://github.com/Zhyhan245/voltapro-service-inquiries (тег v1.0.0). Дати: видачі — 28.10.2026, подання та захисту — 27.11.2026.', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=4)
    add_p(doc, 'Завдання видав: Червоненко П. П.   Дата: 28.10.2026     Завдання прийняла: Жиган Д. О.   Дата: 28.10.2026', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=0)

    # =========================================================================
    # АНОТАЦІЯ ТА ABSTRACT
    # =========================================================================
    add_h1(doc, 'АНОТАЦІЯ')
    add_p(doc, 'У курсовій роботі виконано повний цикл розроблення корпоративного вебсайту із системою керування зверненнями клієнтів для інженерно-сервісного підприємства «ВольтаПро Сервіс» (варіант 5) — від аналізу предметної області й аналогів до проєктування архітектури, реалізації, автоматизованого тестування та оцінювання якості. Метою роботи є підвищення прозорості та керованості процесу приймання й опрацювання звернень за рахунок онлайн-реєстрації, захищеного відстеження статусу клієнтом, рольового доступу персоналу та контролю строків виконання (SLA) у робочих днях.')
    add_p(doc, 'Систему побудовано як модульний вебзастосунок на базі Node.js v26.13.1, Express 4.21.2, реляційної СКБД SQLite (вбудований модуль node:sqlite у режимі WAL із зовнішніми ключами, CHECK-обмеженнями та транзакціями BEGIN IMMEDIATE) і клієнтського інтерфейсу на семантичному HTML5, CSS3, ES6+ JavaScript та локально підключеному Bootstrap 5. Реалізовано каталог послуг із пошуком (включно з власною SQL-функцією uk_lower для кирилиці) та фільтрацією, форму звернення із захистом від повторного надсилання (idempotency_key), пошук заявки клієнтом за номером і 6-символьним кодом доступу без витоку внутрішніх приміток, кабінет менеджера з рольовою моделлю (MANAGER, SENIOR_MANAGER), обов’язковою причиною відхилення та забороною несанкціонованого повернення завершених заявок. Як погоджене ускладнення реалізовано обчислення дедлайнів SLA у робочих днях за явно визначеним виробничим календарем та інтерактивний звіт про прострочені заявки.')
    add_p(doc, 'Працездатність підтверджено автоматизованими модульними та інтеграційними тестами (усі 7 тестових блоків пройдено, 0 помилок, покриття 90,00 % рядків коду), 10 наскрізними перевірками у Google Chrome на ширинах 360, 768 та 1280 пікселів, навантажувальним тестуванням (до 800,99 запитів/с при 25 віртуальних користувачах, p95 = 7,37 мс, 0 % помилок) та тестуванням обсягом даних до 10 000 записів.')
    add_p(doc, 'Ключові слова: корпоративний сайт, керування зверненнями, контроль SLA, робочі дні, Node.js, Express, SQLite, Bootstrap 5, RBAC, RFC 9457.')

    add_h1(doc, 'ABSTRACT')
    add_p(doc, 'This coursework presents the complete development lifecycle of a corporate website with a client inquiry management system for the engineering and service enterprise "VoltaPro Service" (Variant 5), covering domain analysis, competitor comparison, architecture design, implementation, automated verification, and quality evaluation. The objective is to increase transparency and responsiveness in processing customer inquiries through online registration, secure status lookup, role-based staff access control, and business-day SLA tracking.')
    add_p(doc, 'The application is built as a modular client-server web system using Node.js v26.13.1, Express 4.21.2, SQLite (built-in node:sqlite in WAL mode with foreign keys, CHECK constraints, and BEGIN IMMEDIATE transactions), and a responsive frontend powered by semantic HTML5, CSS3, ES6+ JavaScript, and locally hosted Bootstrap 5. Implemented features include a searchable service catalog (with a custom uk_lower SQL function for case-insensitive Cyrillic matching), an idempotent inquiry submission form, public ticket lookup by ticket number and 6-character access code filtering out internal staff notes, a role-based manager dashboard (MANAGER and SENIOR_MANAGER) enforcing mandatory rejection reasons and transition rules, and an agreed extension: business-day SLA deadline calculation using an explicit calendar and an overdue inquiries report.')
    add_p(doc, 'Verification results include automated unit and integration tests (7/7 passed, 0 failures, 90.00% line coverage), 10 end-to-end browser checks in Google Chrome across 360, 768, and 1280 px viewports, load testing up to 800.99 RPS at 25 virtual users (p95 = 7.37 ms, 0% errors), and data-volume benchmarking up to 10,000 inquiry records.')
    add_p(doc, 'Keywords: corporate website, inquiry management, SLA tracking, business days, Node.js, Express, SQLite, Bootstrap 5, RBAC, RFC 9457.')

    # =========================================================================
    # ЗМІСТ ТА ПЕРЕЛІК СКОРОЧЕНЬ
    # =========================================================================
    add_h1(doc, 'ЗМІСТ')
    add_toc_field(doc)

    add_h1(doc, 'ПЕРЕЛІК СКОРОЧЕНЬ ТА УМОВНИХ ПОЗНАЧЕНЬ')
    add_table(
        doc,
        'Перелік основних скорочень та термінів, використаних у пояснювальній записці',
        ['Скорочення', 'Розшифрування та значення у межах роботи'],
        [
            ['AC / ADR', 'Architecture Constraint / Architecture Decision Record — архітектурне обмеження / архітектурне рішення'],
            ['API / JSON', 'Application Programming Interface / JavaScript Object Notation — програмний інтерфейс та формат обміну даними'],
            ['ASVS / OWASP', 'Application Security Verification Standard / Open Worldwide Application Security Project — стандарти безпеки'],
            ['CRUD', 'Create, Read, Update, Delete — базові операції керування даними'],
            ['DTO', 'Data Transfer Object — об’єкт передавання даних, відокремлений від внутрішньої структури таблиць БД'],
            ['ER / SQL', 'Entity-Relationship / Structured Query Language — модель «сутність–зв’язок» та мова структурованих запитів'],
            ['FR / NFR', 'Functional Requirement / Non-Functional Requirement — функціональна та нефункціональна вимоги'],
            ['HTTP / HTTPS', 'HyperText Transfer Protocol (Secure) — протокол передавання гіпертексту'],
            ['IDOR', 'Insecure Direct Object Reference — вразливість прямого доступу до чужого об’єкта'],
            ['LOC', 'Lines of Code — кількість рядків власного програмного коду'],
            ['NIST', 'National Institute of Standards and Technology — автор стандарту автентифікації SP 800-63B-4'],
            ['PK / FK / UK', 'Primary Key / Foreign Key / Unique Key — первинний, зовнішній та унікальний ключі таблиці БД'],
            ['RBAC', 'Role-Based Access Control — керування доступом на основі ролей (MANAGER, SENIOR_MANAGER)'],
            ['REST', 'Representational State Transfer — архітектурний стиль побудови мережевих інтерфейсів'],
            ['RPS / VU', 'Requests Per Second / Virtual Users — запитів за секунду / віртуальні користувачі навантажувального стенда'],
            ['SLA', 'Service Level Agreement — регламентований строк опрацювання звернення у робочих днях'],
            ['TC', 'Test Case — формалізований тестовий сценарій перевірки вимоги'],
            ['UI', 'User Interface — користувацький інтерфейс'],
            ['UML', 'Unified Modeling Language — уніфікована мова моделювання (версія 2.5.1)'],
            ['WAL', 'Write-Ahead Logging — режим журналювання SQLite для паралельного читання та надійних транзакцій'],
            ['WCAG', 'Web Content Accessibility Guidelines — настанови щодо доступності вебконтенту (версія 2.2)'],
            ['XSS', 'Cross-Site Scripting — міжсайтовий скриптинг'],
            ['СКБД / БД', 'Система керування базами даних / база даних']
        ],
        col_widths=[4.2, 12.8]
    )

    # =========================================================================
    # ВСТУП
    # =========================================================================
    add_h1(doc, 'ВСТУП')
    add_p(doc, 'Сучасні інженерно-сервісні підприємства щодня отримують значну кількість звернень від приватних та корпоративних клієнтів щодо енергоаудиту, проєктування, монтажу та технічного обслуговування систем живлення. Коли приймання таких запитів відбувається через розрізнені канали — електронну пошту, телефонні дзвінки та месенджери — заявки нерідко губляться, клієнт не має прозорої інформації про стан розгляду свого звернення, а керівництво не бачить реального навантаження на менеджерів і порушення строків опрацювання.')
    add_p(doc, 'На практиці типові недоліки спрощених вебформ зворотного зв’язку полягають у відсутності єдиного номера відстеження, створенні дублікатів при повторному натисканні кнопки надсилання, збереженні лише клієнтської валідації без серверної перевірки, відсутності розмежування між публічними відповідями клієнту й внутрішніми службовими примітками, а також у підміні робочих днів календарними при контролі строків виконання.')
    add_p(doc, 'Метою курсової роботи є підвищення прозорості та керованості процесу приймання й опрацювання клієнтських звернень підприємства «ВольтаПро Сервіс» шляхом розроблення корпоративного вебсайту з каталогом послуг, ідемпотентною формою звернення, захищеним пошуком статусу заявки, рольовим кабінетом менеджера та контролем строків опрацювання SLA за явно визначеним виробничим календарем.')
    add_p(doc, 'Для досягнення поставленої мети визначено такі завдання:')
    for t in [
        'проаналізувати предметну область опрацювання корпоративних звернень, порівняти щонайменше три релевантні аналоги та сформулювати перевірювані функціональні й нефункціональні вимоги;',
        'розробити й обґрунтувати ключові алгоритми системи — календарний розрахунок дедлайну та прострочення SLA у робочих днях, скінченний автомат допустимих переходів статусів і перевірки ролей, а також алгоритм ідемпотентної реєстрації звернення;',
        'спроєктувати архітектуру вебсистеми у розрізі основних архітектурних подань (UI, Logical, Deployment, Design, Process, Data, Interface, Security, Operational, Infrastructure та Delivery Strategy);',
        'реалізувати серверну частину на Node.js + Express + SQLite із міграціями, синтетичним демонстраційним набором записів та адаптивний клієнтський інтерфейс на HTML5, CSS3, JavaScript і Bootstrap 5;',
        'провести комплексне забезпечення якості: модульне, інтеграційне, наскрізне браузерне (360/768/1280 px), навантажувальне тестування та тестування обсягом даних до 10 000 записів, а також підготувати супровідну документацію.'
    ]:
        add_bullet(doc, t)
    add_p(doc, 'Об’єктом розроблення є процес приймання, розподілу та контролю строків опрацювання клієнтських звернень підприємства. Предметом розроблення є реляційна модель даних, бізнес-правила переходів статусів і доступу, алгоритми календарного контролю SLA, REST API та вебінтерфейс системи «ВольтаПро Сервіс».')
    add_p(doc, 'Навчальний статус і межі системи: підприємство «ВольтаПро Сервіс», перелік його послуг, облікові записи працівників (senior@voltapro.ua, manager1@voltapro.ua, manager2@voltapro.ua) та 18 демонстраційних звернень є синтетичними навчальними даними. Реальне надсилання SMS чи електронних листів винесено за межі роботи.')

    # =========================================================================
    # РОЗДІЛ 1. АНАЛІЗ ПРЕДМЕТНОЇ ОБЛАСТІ ТА ВИМОГИ
    # =========================================================================
    add_h1(doc, '1 АНАЛІЗ ПРЕДМЕТНОЇ ОБЛАСТІ ТА ВИМОГИ')

    add_h2(doc, '1.1 Проблема, цільова аудиторія та аналіз аналогів')
    add_p(doc, 'До впровадження вебсистеми приймання звернень на умовному підприємстві «ВольтаПро Сервіс» відбувалося через загальну поштову скриньку та телефонні дзвінки. У такому процесі виникають системні проблеми (обґрунтовані навчальні припущення): відсутній унікальний реєстраційний номер звернення; клієнти повторно надсилають запити, не маючи змоги перевірити статус онлайн; менеджери можуть помилково переслати внутрішній службовий коментар клієнтові; облік строків виконання ведеться вручну без урахування вихідних і святкових днів.')
    add_p(doc, 'Цільовою аудиторією системи є дві основні групи користувачів: по-перше, зовнішні відвідувачі та клієнти (приватні особи й представники бізнесу), яким потрібен зрозумілий каталог послуг, швидке подання звернення та перевірка його стану; по-друге, менеджери й керівник сервісного відділу, які розподіляють заявки, ведуть службове листування та контролюють прострочення SLA.')
    add_p(doc, 'Для визначення кращих практик виконано порівняльний аналіз трьох релевантних систем керування зверненнями за відкритою документацією (дата ознайомлення — 05.11.2026): Zendesk Support, Bitrix24 CRM та Jira Service Management (таблиця 1.1).')

    add_table(
        doc,
        'Таблиця 1.1 — Порівняльний аналіз релевантних аналогів керування зверненнями',
        ['Критерій', 'Zendesk Support', 'Bitrix24 CRM', 'Jira Service Management', 'Вимога до власного рішення'],
        [
            ['Цільова аудиторія', 'Служби підтримки середнього й великого бізнесу', 'Відділи продажів і малий/середній бізнес', 'ІТ-команди та сервісні підрозділи (ITSM)', 'Мале/середнє підприємство: гість, менеджер, старший менеджер'],
            ['Основний процес', 'Тикет-система з чергами та макросами', 'Воронка лідів з вебформ у CRM', 'Портал запитів із процесами погодження', 'Публічний каталог → форма звернення → пошук статусу → кабінет менеджера + SLA'],
            ['Робота з помилками та приватністю', 'Перевірено за док.: розділення публічних відповідей і internal notes', 'Не перевірено у платних тарифах', 'Перевірено за док.: обов’язкові поля переходів workflow', 'Чіткі помилки RFC 9457, збереження полів форми, суворе фільтрування внутрішніх нотаток'],
            ['Доступність і мобільний інтерфейс', 'Адаптивний вебпортал, підтримка WCAG', 'Адаптивний веб і мобільний застосунок', 'Адаптивний портал замовника', 'Контрольні точки 360, 768, 1280 px; навігація клавіатурою (WCAG 2.2)'],
            ['Обмеження застосування', 'Платна підписка, надлишкова складність', 'Закрита екосистема, прив’язка до хмарних тарифів', 'Високий поріг входу для зовнішніх клієнтів', 'Легкий автономний стек, відкритий код, прозорий календар робочих днів SLA']
        ],
        col_widths=[2.8, 3.5, 3.5, 3.5, 3.7]
    )

    add_p(doc, 'Висновок аналізу аналогів: промислові сервіс-дески перевантажені платними функціями та складними інтеграціями, але задають правильні архітектурні орієнтири — розділення публічних і внутрішніх коментарів (як у Zendesk), обов’язкові поля при відхиленні (як у Jira) та календарний облік SLA. Ці рішення адаптовано в легкому модульному застосунку «ВольтаПро Сервіс».')

    add_h2(doc, '1.2 Типи користувачів, розмежування повноважень і межі системи')
    add_p(doc, 'У системі визначено три ролі користувачів (таблиця 1.2): Гість / Клієнт, Менеджер (MANAGER) та Старший менеджер (SENIOR_MANAGER). Ролі описують набір дозволених операцій і перевіряються на сервері для кожного захищеного запиту та конкретного об’єкта.')

    add_table(
        doc,
        'Таблиця 1.2 — Матриця розмежування повноважень вебсистеми «ВольтаПро Сервіс»',
        ['Операція / маршрут', 'Гість / Клієнт', 'Менеджер (MANAGER)', 'Старший менеджер (SENIOR_MANAGER)', 'Механізм серверної перевірки'],
        [
            ['Каталог послуг (GET /api/services)', 'Так', 'Так', 'Так', 'Публічний маршрут; вибірка лише is_active = 1'],
            ['Подання звернення (POST /api/inquiries)', 'Так', 'Так', 'Так', 'Валідація полів, idempotency_key, розрахунок due_date'],
            ['Пошук заявки (GET /api/inquiries/lookup)', 'Так (за кодом)', 'Так (за кодом)', 'Так (за кодом)', 'Перевірка ticketNumber + accessCode; is_internal = 0'],
            ['Робочий список (GET /api/manager/inquiries)', 'Ні (401)', 'Лише свої та нерозподілені', 'Усі звернення', 'requireAuth + SQL-фільтр за assigned_employee_id'],
            ['Картка заявки (GET /api/manager/inquiries/:id)', 'Ні (401)', 'Лише доступна (403 для чужої)', 'Будь-яка заявка', 'assertEmployeeCanAccessInquiry()'],
            ['Призначення (PATCH .../assign)', 'Ні (401)', 'Лише на себе (403 іншим)', 'На будь-кого або зняття', 'Перевірка ролі в assignInquiry()'],
            ['Перехід у роботу / завершення / відхилення', 'Ні (401)', 'Для доступних заявок', 'Для будь-якої заявки', 'validateStatusTransition() + причина відхилення'],
            ['Повернення завершеної/відхиленої в роботу', 'Ні (401)', 'Ні (403 REOPEN_FORBIDDEN)', 'Так', 'Перевірка ролі SENIOR_MANAGER при термінальному статусі'],
            ['Звіт прострочень SLA (GET .../reports/overdue)', 'Ні (401)', 'За доступними заявками', 'Повний звіт', 'requireAuth + evaluateSlaState()']
        ],
        col_widths=[4.4, 2.3, 3.3, 3.3, 3.7]
    )

    add_p(doc, 'Межі власної розробки охоплюють публічний сайт, кабінет менеджера, REST API, доменні правила, БД SQLite та інфраструктуру запуску. Єдина стороння залежність інтерфейсу — Bootstrap 5 — підключена локально. Зовнішні поштові чи SMS-сервери свідомо винесено за межі роботи.')

    add_h2(doc, '1.3 Функціональні вимоги')
    add_p(doc, 'Функціональні можливості згруповано у шість блоків (F-01 — каталог; F-02 — подання звернення; F-03 — публічний пошук; F-04 — керування статусами; F-05 — розмежування доступу та внутрішні нотатки; F-06 — календар SLA і звіт прострочень) та деталізовано в таблиці 1.3. Паспорти двох критичних вимог наведено в таблицях 1.4 і 1.5.')

    add_table(
        doc,
        'Таблиця 1.3 — Каталог функціональних вимог вебсистеми',
        ['Код', 'Функція', 'Стисле формулювання перевірюваної вимоги'],
        [
            ['FR-01', 'F-01 Каталог', 'Пошук за назвою, описом і категорією (у т.ч. кирилиця без урахування регістру) та фільтрація за категорією активних послуг.'],
            ['FR-02', 'F-02 Звернення', 'Серверна валідація контактних даних (неповний контакт → 400 VALIDATION_ERROR із fieldErrors) та збереження даних у формі інтерфейсу.'],
            ['FR-03', 'F-02 Ідемпотентність', 'При повторному надсиланні форми з тим самим idempotencyKey дублікат у БД не створюється; повертається 200 (duplicate: true).'],
            ['FR-04', 'F-03 Пошук', 'Пошук заявки за номером VP-YYYY-NNNN та 6-символьним кодом доступу; при помилці коду — 404 без розкриття деталей.'],
            ['FR-05', 'F-05 Приватність', 'Публічний пошук повертає лише публічну відповідь та історію з is_internal = 0; внутрішні примітки ніколи не потрапляють клієнту.'],
            ['FR-06', 'F-04 Списки й призначення', 'Робочий список менеджера з фільтрами, пошуком, пагінацією та призначенням відповідального (MANAGER — лише на себе).'],
            ['FR-07', 'F-04 Переходи статусів', 'Допустимі переходи «NEW → IN_PROGRESS → COMPLETED / REJECTED»; при відхиленні причина (≥ 5 симв.) обов’язкова (400).'],
            ['FR-08', 'F-05 Захист об’єкта', 'Заборона доступу MANAGER до чужої заявки (403) та повернення завершеної/відхиленої заявки в роботу без ролі SENIOR_MANAGER (403).'],
            ['FR-09', 'F-06 Календар SLA', 'Автоматичний розрахунок due_date у робочих днях (пропуск сб/нд та свят із calendar_exceptions) при створенні звернення.'],
            ['FR-10', 'F-06 Звіт SLA', 'Формування звіту прострочених заявок із підрахунком прострочених робочих днів та відображенням календаря винятків.']
        ],
        col_widths=[1.8, 3.8, 11.4]
    )

    add_table(
        doc,
        'Таблиця 1.4 — Паспорт функціональної вимоги FR-02 / FR-03 (реєстрація звернення)',
        ['Поле', 'Зміст'],
        [
            ['Ідентифікатор', 'FR-02 / FR-03'],
            ['Функція', 'Приймання клієнтського звернення із захистом від дублікатів'],
            ['Актор', 'Гість / Клієнт'],
            ['Передумови', 'Послуга існує та є активною (is_active = 1)'],
            ['Вхід', 'serviceId, contactName, companyName, phone, email, subject, message, idempotencyKey'],
            ['Основний результат', 'Створено 1 контакт, 1 звернення (NEW), запис історії; повернено 201 Created, ticketNumber, accessCode, dueDate'],
            ['Повторний запит', 'З тим самим idempotencyKey повертається 200 OK (duplicate: true) без створення другого запису'],
            ['Некоректні дані', 'Повертається 400 VALIDATION_ERROR (RFC 9457) з масивом fieldErrors; стан БД не змінюється, поля форми зберігаються'],
            ['Критерії приймання', 'Перевірено в tests/integration.test.js та E2E Chrome (рис. 4.3, 4.4)']
        ],
        col_widths=[4.2, 12.8]
    )

    add_table(
        doc,
        'Таблиця 1.5 — Паспорт функціональної вимоги FR-07 / FR-08 (переходи статусів і права)',
        ['Поле', 'Зміст'],
        [
            ['Ідентифікатор', 'FR-07 / FR-08'],
            ['Функція', 'Керування життєвим циклом статусу звернення та перевірка прав ролі'],
            ['Актор', 'Менеджер (MANAGER) або Старший менеджер (SENIOR_MANAGER)'],
            ['Передумови', 'Працівник автентифікований; для MANAGER заявка нерозподілена або призначена йому'],
            ['Вхід', 'inquiryId, status (IN_PROGRESS | COMPLETED | REJECTED), rejectionReason, publicResponse'],
            ['Основний результат', 'Статус оновлено у транзакції; зафіксовано closed_at (для фінальних станів) і запис в inquiry_history'],
            ['Відмови та конфлікти', 'Непорожня причина відсутня при REJECTED → 400; перехід COMPLETED→IN_PROGRESS для MANAGER → 403; недозволений перехід → 409'],
            ['Критерії приймання', 'Перевірено модульними тестами (unit.test.js) та інтеграційним тестом (integration.test.js)']
        ],
        col_widths=[4.2, 12.8]
    )

    add_h2(doc, '1.4 Нефункціональні вимоги та трасування')
    add_p(doc, 'Нефункціональні вимоги (таблиця 1.6) визначають вимірювані критерії якості системи. Матрицю трасування винесено в Додаток А.')
    add_table(
        doc,
        'Таблиця 1.6 — Перевірювані нефункціональні вимоги вебсистеми',
        ['Код', 'Категорія', 'Формулювання вимоги', 'Спосіб перевірки'],
        [
            ['NFR-SEC-01', 'Безпека', 'Паролі хешуються scrypt із сіллю (NIST SP 800-63B-4, ≥ 15 симв.); сесійні cookie мають HttpOnly та SameSite=Lax; чужі заявки повертають 403; параметризовані SQL-запити.', 'unit.test.js, integration.test.js, аудит коду.'],
            ['NFR-DATA-01', 'Цілісність даних', 'Схема БД відтворюється міграцією 001_init.sql; перезапуск сервера не втрачає звернення; зовнішні ключі й CHECK-обмеження діють на рівні SQLite.', 'migrate.js, seed.js, перезапуск сервера.'],
            ['NFR-UI-01', 'Адаптивність і доступність', 'Відсутність горизонтального переповнення на 360, 768 і 1280 px; видимий фокус (:focus-visible), skip-link, текстове дублювання статусів (WCAG 2.2).', 'e2e-screenshots.js у Google Chrome + візуальна перевірка.'],
            ['NFR-PERF-01', 'Продуктивність', 'На тестовому стенді (5, 10, 25 VU) p95 часу відповіді API ≤ 800 мс, помилок < 1 %.', 'load-test.js та volume-test.js (до 10 000 записів).'],
            ['NFR-RUN-01', 'Відтворюваність', 'Проєкт запускається за README.md у чистому каталозі без секретів у коді (або через Docker Compose).', 'Контрольний запуск npm install → migrate → seed → start.']
        ],
        col_widths=[2.5, 2.8, 7.2, 4.5]
    )

    add_h2(doc, '1.5 Моделювання прецедентів')
    add_p(doc, 'Взаємодію трьох ролей користувачів із системою подано на UML-діаграмі прецедентів [15] (рисунок 1.1).')
    add_figure(doc, DIAG_DIR / 'fig_1_1_use_cases.png', 'Рисунок 1.1 — UML-діаграма прецедентів користувачів вебсистеми «ВольтаПро Сервіс»', width_cm=15.0)
    add_p(doc, 'Специфікація ключового прецеденту «Подання звернення»: актор — Гість / Клієнт; передумови — активна послуга в каталозі; основний успішний результат — створено контакт, звернення зі статусом NEW, обчислено дедлайн SLA в робочих днях, видано номер і код доступу; альтернативи — при помилці валідації повертаються підказки зі збереженням полів, при повторному натисканні повертається наявний запис; мінімальна гарантія — відкат транзакції без часткового збереження.')

    add_h2(doc, '1.6 Перевірка якості вимог')
    add_table(
        doc,
        'Таблиця 1.7 — Результати перевірки якості формулювання вимог',
        ['Властивість', 'Критерій перевірки', 'Підтвердження для проєкту «ВольтаПро Сервіс»'],
        [
            ['Атомарність', 'Окрема вимога описує одне правило', 'Розділено валідацію (FR-02), ідемпотентність (FR-03), приватність приміток (FR-05), переходи (FR-07) і SLA (FR-09).'],
            ['Перевірюваність', 'Задано вхід, дію та точний результат', 'Кожна вимога має HTTP-статус, код помилки RFC 9457 і прив’язаний тест.'],
            ['Зрозумілість', 'Єдина термінологія ролей і станів', 'Узгоджено: NEW, IN_PROGRESS, COMPLETED, REJECTED; MANAGER, SENIOR_MANAGER.'],
            ['Повнота', 'Охоплено успіх, відмови й межі', 'Описано неповний контакт, дублювання, невірний код пошуку, чужу заявку, відхилення без причини.'],
            ['Несуперечливість', 'Відсутність конфліктів правил', 'Усі обмеження статусів і ролей перевіряються централізовано в src/domain/inquiryRules.js.'],
            ['Простежуваність', 'Зв’язок коду вимоги з тестом і доказом', 'Реалізовано в матриці трасування (Додаток А, таблиця А.1).']
        ],
        col_widths=[3.2, 5.2, 8.6]
    )

    # =========================================================================
    # РОЗДІЛ 2. АЛГОРИТМІЧНЕ ЗАБЕЗПЕЧЕННЯ ВЕБСИСТЕМИ
    # =========================================================================
    add_h1(doc, '2 АЛГОРИТМІЧНЕ ЗАБЕЗПЕЧЕННЯ ВЕБСИСТЕМИ')

    add_h2(doc, '2.1 Ключові алгоритми та їхні паспорти')
    add_p(doc, 'Поведінку системи визначають чотири ключові алгоритми: 1) розрахунок дати дедлайну SLA у робочих днях (addBusinessDays); 2) оцінка стану SLA та підрахунок прострочених робочих днів (evaluateSlaState); 3) перевірка допустимості переходу статусу та прав ролі (validateStatusTransition); 4) ідемпотентна транзакційна реєстрація звернення (createPublicInquiry). Паспорти алгоритмів наведено в таблиці 2.1.')

    add_table(
        doc,
        'Таблиця 2.1 — Паспорти ключових алгоритмів вебсистеми',
        ['Алгоритм', 'Призначення (вимоги)', 'Вхід / Вихід', 'Інваріант та обробка відмов', 'Складність'],
        [
            ['1. addBusinessDays', 'FR-09: обчислення дедлайну SLA за календарем', 'Вхід: startDate, slaDays (1..30), exceptionsMap. Вихід: dueDate (YYYY-MM-DD).', 'Повернена дата завжди є робочим днем; між startDate і dueDate рівно slaDays робочих днів. Ліміт guard = 3660.', 'Час: O(d), d ≤ 30; пам’ять: O(1).'],
            ['2. evaluateSlaState', 'FR-10: визначення стану та прострочення', 'Вхід: inquiry, referenceDate, exceptionsMap. Вихід: {state, overdueBusinessDays, remainingBusinessDays}.', 'Розрізняє ON_TIME, DUE_TODAY, OVERDUE, CLOSED_ON_TIME, CLOSED_OVERDUE. Виключає вихідні та свята.', 'Час: O(Δt) днів; пам’ять: O(1).'],
            ['3. validateStatusTransition', 'FR-07, FR-08: контроль життєвого циклу', 'Вхід: currentStatus, nextStatus, employeeRole, rejectionReason. Вихід: reason або DomainError.', 'Забороняє недозволені переходи (409), повернення завершених без SENIOR_MANAGER (403), відмову без причини (400).', 'Час: O(1); пам’ять: O(1).'],
            ['4. createPublicInquiry', 'FR-02, FR-03: атомарне створення без дублів', 'Вхід: body (DTO). Вихід: {duplicate, ticketNumber, accessCode, dueDate}.', 'Транзакція BEGIN IMMEDIATE; унікальний idempotency_key виключає дублікати при повторному запиті.', 'Час: O(log n + d); пам’ять: O(1).']
        ],
        col_widths=[3.2, 3.2, 3.8, 4.3, 2.5]
    )

    add_p(doc, 'Математична модель обчислення кількості прострочених робочих днів для активної заявки на звітну дату t_ref (за умови t_ref > t_due) задається формулою:')
    add_p(doc, 'K_overdue(t_due, t_ref) = Σ I_work(t_i),   де t_i ∈ (t_due; t_ref]', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=4)
    add_p(doc, 'де t_due — планова дата дедлайну (YYYY-MM-DD); t_ref — поточна або звітна дата; I_work(t_i) ∈ {0, 1} — індикаторна функція робочого дня, яка повертає 1, якщо день t_i є робочим з урахуванням таблиці винятків calendar_exceptions (інакше для суботи й неділі повертає 0, а для буднів — 1).')

    add_h2(doc, '2.2 Псевдокод та блок-схеми ключових алгоритмів')
    add_code_block(
        doc,
        'Лістинг 2.1 — Псевдокод алгоритму addBusinessDays (src/domain/calendarRules.js)',
        """addBusinessDays(startDateStr, businessDays, exceptionsMap):
  якщо businessDays < 1 або не ціле -> виняток
  cursor = toIsoDateOnly(startDateStr)
  remaining = businessDays
  поки remaining > 0:
    cursor = addCalendarDays(cursor, 1)
    якщо isWorkingDay(cursor, exceptionsMap):
      remaining = remaining - 1
  повернути cursor"""
    )
    add_figure(doc, DIAG_DIR / 'fig_2_1_sla_algorithm.png', 'Рисунок 2.1 — Блок-схема алгоритму розрахунку дедлайну SLA у робочих днях', width_cm=14.5)
    add_figure(doc, DIAG_DIR / 'fig_2_2_status_transition.png', 'Рисунок 2.2 — Блок-схема перевірки допустимості переходу статусу звернення', width_cm=14.5)

    add_h2(doc, '2.3 Оцінка складності та експериментальне порівняння')
    add_p(doc, 'Для алгоритму addBusinessDays при обмеженні d ≤ 30 кількість ітерацій циклу обмежена зверху константою (приблизно 45 ітерацій з урахуванням вихідних), тому практична складність є O(1) за часом і O(1) за пам’яттю. Для звіту прострочень SLA (getOverdueReport) нехай n — кількість активних звернень: вибірка з БД використовує складений індекс idx_inquiries_status_due(status, due_date), оцінка кожного запису потребує O(1) операцій у межах робочого горизонту, а сортування за спаданням прострочених днів стандартним TimSort у V8 виконується за O(n log n) часу та O(n) пам’яті.')
    add_p(doc, 'Експериментальна перевірка на реальному стенді (таблиця 2.2, детальніше в розділі 5.7) показала, що при зростанні БД у 100 разів (від 100 до 10 000 записів) середній час виконання фільтрованого списку зріс із 0,295 мс до 11,297 мс, а звіту прострочень SLA — із 0,527 мс до 53,691 мс, що повністю підтверджує прийнятну лінійно-логарифмічну масштабованість рішення.')

    # =========================================================================
    # РОЗДІЛ 3. ПРОЄКТУВАННЯ ВЕБСИСТЕМИ
    # =========================================================================
    add_h1(doc, '3 ПРОЄКТУВАННЯ ВЕБСИСТЕМИ')

    add_h2(doc, '3.1 Архетип і обмеження архітектури (Architecture Constraints)')
    add_p(doc, 'Вебсистема «ВольтаПро Сервіс» реалізована як модульний клієнт-серверний моноліт: клієнтська частина (HTML5 + CSS3 + ES6+ JavaScript + Bootstrap 5) обмінюється JSON-повідомленнями через REST API із сервером Node.js + Express, який керує локальним сховищем SQLite. Архітектура модульного моноліту обрана замість мікросервісів для забезпечення атомарних транзакцій у межах єдиної БД, простоти розгортання та низьких витрат супроводу.')
    add_table(
        doc,
        'Таблиця 3.1 — Паспорт архітектурних обмежень вебсистеми',
        ['Код і зміст', 'Джерело обмеження', 'Наслідок і спосіб перевірки'],
        [
            ['AC-01: HTML5, CSS3, JS, Bootstrap 5 без важких SPA-фреймворків', 'Побажання та узгодження стека', 'Модулі ES6 у public/js/, локальна роздача Bootstrap 5 без залежності від інтернету.'],
            ['AC-02: Node.js + Express + вбудований node:sqlite', 'Вибір технологій (табл. 1.3)', 'Відсутність проблем із нативною компіляцією; перевірено на Node v26.13.1.'],
            ['AC-03: Відсутність платних сервісів і реальних розсилок', 'Межі курсової роботи', 'Уся робота виконується локально; історія дій зберігається в БД.'],
            ['AC-04: Підтримка Docker Compose', 'Вимоги до розгортання', 'Файли Dockerfile та docker-compose.yml із томом ./data.']
        ],
        col_widths=[5.5, 4.5, 7.0]
    )

    add_h2(doc, '3.2 Подання користувацького інтерфейсу (UI View)')
    add_p(doc, 'Інтерфейс складається з трьох сторінок (/, /status, /manager) та восьми станів (рисунок 3.1, таблиця 3.2). Передбачено стани завантаження, порожньої вибірки, помилки валідації зі збереженням введених полів, успіху та відмови 401/403. Адаптивність перевіряється на ширинах 360, 768 і 1280 CSS-пікселів.')
    add_figure(doc, DIAG_DIR / 'fig_3_1_ui_state_map.png', 'Рисунок 3.1 — Карта екранів та станів користувацького інтерфейсу (UI View)', width_cm=15.5)
    add_table(
        doc,
        'Таблиця 3.2 — Паспорт екранів інтерфейсу та їхніх станів',
        ['Екран', 'Доступні ролі', 'Основні дані, стани та UI-патерни'],
        [
            ['UI-01: Каталог послуг (/)', 'Гість, Менеджер', 'GET /api/services; пошук, фільтр категорій, картки послуг, порожня вибірка.'],
            ['UI-02 / UI-03: Форма та підтвердження (/)', 'Гість, Менеджер', 'POST /api/inquiries; інлайн-валідація, блокування кнопки, збереження полів при 400, картка номера й коду.'],
            ['UI-04: Перевірка статусу (/status)', 'Гість / Клієнт', 'GET /api/inquiries/lookup; статус, дедлайн SLA, публічна відповідь без внутрішніх нотаток.'],
            ['UI-05: Вхід менеджера (/manager)', 'Персонал', 'POST /api/auth/login; форма входу, помилка 401, демо-підказка обліковок.'],
            ['UI-06 / UI-07 / UI-08: Кабінет, картка, звіт SLA (/manager)', 'MANAGER, SENIOR_MANAGER', 'Фільтрована таблиця з пагінацією, модальна картка з історією та примітками, вкладка звіту SLA.']
        ],
        col_widths=[4.6, 3.4, 9.0]
    )

    add_h2(doc, '3.3 Логічне подання системи (Logical View)')
    add_p(doc, 'Логічне подання (рисунок 3.2) ізолює клієнтський код у браузері від приватного сховища SQLite через довірений сервер Node.js + Express.')
    add_figure(doc, DIAG_DIR / 'fig_3_2_logical_view.png', 'Рисунок 3.2 — Структурна схема вебсистеми та її довіреної частини (Logical View)', width_cm=15.0)

    add_h2(doc, '3.4 Подання розгортання (Deployment View)')
    add_p(doc, 'Схема розгортання (рисунок 3.3, таблиця 3.3) показує взаємодію браузера Google Chrome на Windows 11 Pro із сервером Node.js (порт 3000) та файловим сховищем SQLite data/voltapro.db (або через Docker Compose).')
    add_figure(doc, DIAG_DIR / 'fig_3_3_deployment_view.png', 'Рисунок 3.3 — Схема розгортання вебсистеми «ВольтаПро Сервіс» (Deployment View)', width_cm=14.5)
    add_table(
        doc,
        'Таблиця 3.3 — Опис вузлів розгортання',
        ['Вузол', 'Параметри середовища і конфігурації', 'Доказ відтворюваності'],
        [
            ['Клієнт', 'Windows 11 Pro, Google Chrome, ширини 360/768/1280 px, доступ до http://localhost:3000', 'e2e-screenshots.js (10/10 перевірок)'],
            ['Вебсервер', 'Node.js v26.13.1, Express 4.21.2, порт 3000, конфігурація в .env', 'npm start + GET /api/health → 200 OK'],
            ['Сховище', 'SQLite (node:sqlite) у WAL-режимі, файл data/voltapro.db, міграція 001_init.sql', 'npm run migrate + ідемпотентний npm run seed']
        ],
        col_widths=[2.8, 7.6, 6.6]
    )

    add_h2(doc, '3.5 Подання шарів і компонентів (Design View)')
    add_p(doc, 'Серверну частину поділено на 4 логічні шари (рисунок 3.4): Подання (routes, middleware), Прикладний шар (services), Предметні правила (domain) та Інфраструктурні адаптери (repositories, db, logger).')
    add_figure(doc, DIAG_DIR / 'fig_3_4_design_layers.png', 'Рисунок 3.4 — Розподіл відповідальності між логічними шарами (Design View)', width_cm=14.5)

    add_h2(doc, '3.6 Подання процесів (Process View)')
    add_p(doc, 'На діаграмі послідовності (рисунок 3.5) показано процес створення звернення з трьома гілками: ідемпотентним повтором, успішним збереженням у транзакції BEGIN IMMEDIATE та відмовою 400 Bad Request при помилці валідації.')
    add_figure(doc, DIAG_DIR / 'fig_3_5_sequence_create.png', 'Рисунок 3.5 — Діаграма послідовності реєстрації звернення (Process View)', width_cm=15.5)

    add_h2(doc, '3.7 Подання даних (Data View)')
    add_p(doc, 'Модель даних у SQLite містить 7 таблиць (рисунок 3.6): services, client_contacts, inquiries, inquiry_history, employees, sessions, calendar_exceptions. Доменні сутності та структури DTO розмежовано (рисунок 3.7), а словник даних основної таблиці inquiries наведено в таблиці 3.4.')
    add_figure(doc, DIAG_DIR / 'fig_3_6_er_model.png', 'Рисунок 3.6 — ER-модель реляційної бази даних SQLite (Data View)', width_cm=15.5)
    add_figure(doc, DIAG_DIR / 'fig_3_7_domain_dto.png', 'Рисунок 3.7 — Діаграма основних доменних об’єктів та DTO', width_cm=15.5)
    add_table(
        doc,
        'Таблиця 3.4 — Словник даних основної сутності inquiries',
        ['Поле', 'Тип / обмеження', 'Призначення'],
        [
            ['id', 'INTEGER PRIMARY KEY AUTOINCREMENT', 'Внутрішній сурогатний ідентифікатор'],
            ['ticket_number', 'TEXT NOT NULL UNIQUE', 'Публічний номер формату VP-YYYY-NNNN'],
            ['access_code', 'TEXT NOT NULL', '6-символьний код доступу для пошуку клієнтом'],
            ['service_id / contact_id', 'INTEGER NOT NULL FK', 'Зовнішні ключі на послугу та контакт клієнта'],
            ['assigned_employee_id', 'INTEGER FK (NULLABLE)', 'Відповідальний менеджер'],
            ['status', 'TEXT CHECK IN (...)', 'NEW, IN_PROGRESS, COMPLETED, REJECTED'],
            ['rejection_reason', 'TEXT (NULLABLE)', 'Обов’язкова причина відхилення'],
            ['idempotency_key', 'TEXT UNIQUE', 'Ключ захисту від повторного надсилання'],
            ['sla_business_days / due_date', 'INTEGER / TEXT', 'Зафіксований норматив SLA та дедлайн за календарем'],
            ['created_at / updated_at / closed_at', 'TEXT (ISO 8601)', 'Часові мітки життєвого циклу']
        ],
        col_widths=[4.8, 5.2, 7.0]
    )

    add_h2(doc, '3.8 Подання інтерфейсів і контрактів (Interface View)')
    add_p(doc, 'Усі програмні інтерфейси описано у специфікації OpenAPI 3.1 (docs/openapi.yaml). Основні маршрути подано в таблиці 3.5, а формат помилок відповідає RFC 9457 Problem Details.')
    add_table(
        doc,
        'Таблиця 3.5 — Каталог основних HTTP-операцій REST API',
        ['Метод і маршрут', 'Доступ', 'Призначення та типові відповіді'],
        [
            ['GET /api/health', 'Публічний', 'Перевірка стану процесу та SQLite (200 / 503)'],
            ['GET /api/services', 'Публічний', 'Каталог послуг із пошуком і фільтрами (200)'],
            ['POST /api/inquiries', 'Публічний', 'Реєстрація звернення з ідемпотентністю (201 / 200 / 400 / 404)'],
            ['GET /api/inquiries/lookup', 'Публічний (за кодом)', 'Пошук за номером і кодом без внутрішніх нотаток (200 / 400 / 404)'],
            ['POST /api/auth/login | logout', 'Персонал', 'Вхід за scrypt-паролем, HttpOnly cookie (200 / 401 | 204)'],
            ['GET /api/manager/inquiries', 'MANAGER / SENIOR', 'Робочий список із фільтрами та пагінацією (200 / 401)'],
            ['PATCH /api/manager/inquiries/:id/assign', 'MANAGER / SENIOR', 'Призначення відповідального з перевіркою ролі (200 / 403)'],
            ['PATCH /api/manager/inquiries/:id/status', 'MANAGER / SENIOR', 'Зміна статусу (200 / 400 / 403 / 409)'],
            ['POST /api/manager/inquiries/:id/comments', 'MANAGER / SENIOR', 'Публічна відповідь або внутрішня примітка (201 / 400 / 403)'],
            ['GET /api/manager/reports/overdue', 'MANAGER / SENIOR', 'Звіт прострочень SLA та винятки календаря (200 / 401)']
        ],
        col_widths=[6.0, 3.6, 7.4]
    )

    add_h2(doc, '3.9 Подання безпеки (Security View)')
    add_p(doc, 'Модель загроз побудовано за OWASP Top 10:2025: 1) автентифікація — scrypt (N=16384, r=8, p=1) з 16-байтною сіллю, перевірка timingSafeEqual, пароль від 15 символів (NIST SP 800-63B-4); 2) сесії — серверна таблиця sessions, HttpOnly і SameSite=Lax cookie; 3) авторизація — перевірка власника/ролі на кожному об’єкті в assertEmployeeCanAccessInquiry; 4) протидія ін’єкціям — параметризовані SQL-запити, escapeHtml на клієнті, заголовки Helmet.')

    add_h2(doc, '3.10 Операційне та інфраструктурне подання (Operational & Infrastructure Views)')
    add_p(doc, 'Операційний моніторинг реалізовано через структуровані JSON-логи з requestId (без паролів і токенів) та маршрут /api/health, який перевіряє реальний запит SELECT 1 до SQLite. Усі експерименти виконано на стенді: ОС Microsoft Windows 11 Pro, процесор AMD Ryzen 7 9700X (8 ядер, 16 потоків), ОЗП 61,61 ГБ, Node.js v26.13.1, npm 11.20.0, Google Chrome.')

    add_h2(doc, '3.11 Стратегія доставки та технологічний стек (Delivery Strategy View)')
    add_p(doc, 'Продукт передається через Git-репозиторій (https://github.com/Zhyhan245/voltapro-service-inquiries, тег v1.0.0) з інструкцією README.md, .env.example, міграцією, seed-скриптом, тестами, Dockerfile та docker-compose.yml.')

    # =========================================================================
    # РОЗДІЛ 4. РЕАЛІЗАЦІЯ ВЕБСИСТЕМИ
    # =========================================================================
    add_h1(doc, '4 РЕАЛІЗАЦІЯ ВЕБСИСТЕМИ')

    add_h2(doc, '4.1 Ключові модулі, класи та інтерфейси')
    add_p(doc, 'Структуру взаємодії серверних модулів показано на рисунку 4.1, а приклад транзакційного створення звернення — у лістингу 4.1.')
    add_figure(doc, DIAG_DIR / 'fig_4_1_architecture_classes.png', 'Рисунок 4.1 — Ключові модулі та контракти серверної реалізації', width_cm=14.0)
    add_code_block(
        doc,
        'Лістинг 4.1 — Транзакційне створення звернення (src/repositories/inquiryRepository.js)',
        """function createWithContactAndHistory({ contact, inquiry, historyComment }) {
  return runInTransaction(db, () => {
    const contactRes = db.prepare(
      `INSERT INTO client_contacts (contact_name, company_name, phone, email, created_at)
       VALUES (?, ?, ?, ?, ?)`
    ).run(contact.contactName, contact.companyName, contact.phone, contact.email, inquiry.createdAt);
    const contactId = Number(contactRes.lastInsertRowid);
    // Вставка звернення, генерація номера VP-YYYY-NNNN та запис в inquiry_history
    return findById(inquiryId);
  });
}"""
    )

    add_h2(doc, '4.2 Реалізація інтерфейсу та роботи з даними')
    add_p(doc, 'На рисунках 4.2–4.8 наведено реальні скриншоти інтерфейсу, зняті в Google Chrome: публічний сайт, помилку валідації зі збереженням даних, підтвердження прийняття, публічну перевірку статусу, робочий список менеджера, картку з внутрішніми примітками та звіт SLA.')
    add_figure(doc, SCR_DIR / '01-public-desktop-1280.png', 'Рисунок 4.2 — Публічний сайт, каталог послуг та форма звернення (1280 px)', width_cm=14.5)
    add_figure(doc, SCR_DIR / '03-public-form-error-validation.png', 'Рисунок 4.3 — Помилка валідації при неповному контакті зі збереженням полів', width_cm=14.5)
    add_figure(doc, SCR_DIR / '04-public-inquiry-confirmed.png', 'Рисунок 4.4 — Картка підтвердження прийняття звернення з кодом доступу та дедлайном SLA', width_cm=14.5)
    add_figure(doc, SCR_DIR / '05-status-lookup-public.png', 'Рисунок 4.5 — Сторінка публічної перевірки статусу без витоку внутрішніх приміток', width_cm=14.5)
    add_figure(doc, SCR_DIR / '07-manager-list-1280.png', 'Рисунок 4.6 — Робочий список звернень у кабінеті менеджера', width_cm=14.5)
    add_figure(doc, SCR_DIR / '08-manager-modal-detail.png', 'Рисунок 4.7 — Картка звернення з призначенням, зміною статусу та внутрішніми примітками', width_cm=14.5)
    add_figure(doc, SCR_DIR / '09-manager-sla-report.png', 'Рисунок 4.8 — Звіт про прострочені заявки та винятки виробничого календаря', width_cm=14.5)

    add_h2(doc, '4.3 Керування кодом і версіями')
    add_p(doc, 'Вихідний код розміщено у GitHub-репозиторії https://github.com/Zhyhan245/voltapro-service-inquiries під ліцензією MIT (гілка main, релізний тег v1.0.0). Секрети (.env) та node_modules/ виключені через .gitignore.')

    add_h2(doc, '4.4 Метрики та перевірка якості реалізації')
    s = code_rep['summary']
    lf = s['largestFile']
    add_table(
        doc,
        'Таблиця 4.1 — Метрики власного програмного коду (tests/code-metrics.js)',
        ['Показник', 'Виміряне значення', 'Інтерпретація'],
        [
            ['Кількість власних JS-файлів', str(s['filesCount']), 'Модульний поділ на шари (сервер + клієнт)'],
            ['Загальний обсяг коду (рядків)', f"{s['totalLines']} (непорожніх: {s['nonEmptyLines']})", 'Чистий код без згенерованих шаблонів'],
            ['Кількість функцій', str(s['totalFunctions']), 'Середній розмір функції ~19 рядків'],
            ['Найбільший файл', f"{lf['file']} ({lf['totalLines']} рядків)", 'Нижче порогу 500 рядків; містить DTO-мапери'],
            ['Покриття тестами (node:test)', '90,00 % рядків / 88,89 % функцій', 'Перевищує навчальну ціль 70 %'],
            ['Стан залежностей', '0 vulnerabilities (npm audit)', 'Усі 74 пакети без відомих вразливостей']
        ],
        col_widths=[5.5, 4.5, 7.0]
    )

    # =========================================================================
    # РОЗДІЛ 5. ЗАБЕЗПЕЧЕННЯ ЯКОСТІ ВЕБСИСТЕМИ
    # =========================================================================
    add_h1(doc, '5 ЗАБЕЗПЕЧЕННЯ ЯКОСТІ ВЕБСИСТЕМИ')

    add_h2(doc, '5.1 План тестування і зв’язок із вимогами')
    add_p(doc, 'Програма перевірки якості охоплює п’ять рівнів: модульне (node:test), інтеграційне (HTTP API + SQLite), наскрізне браузерне (Playwright + Google Chrome), навантажувальне (5, 10, 25 VU) та тестування обсягом даних (100, 1 000, 10 000 записів).')

    add_h2(doc, '5.2 Тест-кейси та протокол приймання')
    add_table(
        doc,
        'Таблиця 5.1 — Реєстр ключових тест-кейсів та фактичних результатів',
        ['Код', 'Вимога', 'Сценарій і тестові дані', 'Очікуваний / Фактичний результат', 'Статус'],
        [
            ['TC-01', 'FR-01', 'Пошук «монтаж» і фільтр за категорією «Аудит»', '200 OK, коректний кириличний пошук; порожній стан без збігів', 'Пройдено'],
            ['TC-02', 'FR-02', 'Форма з неповним контактом (ім’я «А», тел. «123»)', '400 VALIDATION_ERROR, поля форми збережені', 'Пройдено'],
            ['TC-03', 'FR-03', 'Повторне надсилання з тим самим idempotencyKey', '200 OK (duplicate: true), дублікат у БД не створено', 'Пройдено'],
            ['TC-04', 'FR-04, FR-05', 'Пошук VP-2026-0001 / DEMO01 та невірний код ZZZZ99', '200 OK без внутрішніх приміток; 404 при невірному коді', 'Пройдено'],
            ['TC-05', 'NFR-SEC-01', 'Запит без сесії до /api/manager/inquiries', '401 Unauthorized', 'Пройдено'],
            ['TC-06', 'FR-08', 'Спроба manager2 відкрити чужу заявку manager1', '403 Forbidden, чужі дані не розкрито', 'Пройдено'],
            ['TC-07', 'FR-07', 'Відхилення заявки без причини (порожній рядок)', '400 REJECTION_REASON_REQUIRED', 'Пройдено'],
            ['TC-08', 'FR-08', 'Повернення COMPLETED → IN_PROGRESS ролями MANAGER та SENIOR_MANAGER', '403 REOPEN_FORBIDDEN для MANAGER; 200 OK для SENIOR_MANAGER', 'Пройдено'],
            ['TC-09', 'FR-09', 'Розрахунок SLA: 2026-12-24 + 2 роб. дн. (свято 25.12 + сб/нд)', 'Отримано точну дату 2026-12-29', 'Пройдено'],
            ['TC-10', 'FR-10', 'Звіт прострочень /api/manager/reports/overdue', 'Точні лічильники активних, прострочених і «дедлайн сьогодні»', 'Пройдено']
        ],
        col_widths=[1.5, 2.0, 5.2, 6.3, 2.0]
    )

    add_h2(doc, '5.3 Модульні, інтеграційні та наскрізні перевірки')
    add_p(doc, 'Усі 7 блоків модульних та інтеграційних тестів (node --test --experimental-test-coverage) пройдено без помилок із покриттям 90,00 % рядків та 88,89 % функцій. Усі 10 наскрізних сценаріїв у Google Chrome (tests/e2e-screenshots.js) також завершилися успішно.')

    add_h2(doc, '5.4 Адаптивність, доступність, браузерна сумісність та швидкодія інтерфейсу')
    add_p(doc, 'Перевірка у Google Chrome на ширинах 360, 768 та 1280 CSS-пікселів (рисунки 5.1 та 5.2) підтвердила відсутність горизонтального переповнення (scrollWidth <= clientWidth). Час DOMContentLoaded на локальному стенді становить 30–60 мс.')
    add_figure(doc, SCR_DIR / '02-public-mobile-360.png', 'Рисунок 5.1 — Мобільна адаптивність публічного сайту на ширині 360 px', width_cm=8.0)
    add_figure(doc, SCR_DIR / '10-manager-tablet-768.png', 'Рисунок 5.2 — Планшетна адаптивність кабінету менеджера на ширині 768 px', width_cm=12.5)

    add_h2(doc, '5.5 Навантажувальне тестування та тестування обсягом даних')
    add_p(doc, 'Результати навантажувального прогону (tests/load-test.js) на 5, 10 та 25 віртуальних користувачах наведено в таблиці 5.2 та на рисунку 5.3; результати масштабування БД (100, 1 000, 10 000 записів, tests/volume-test.js) — у таблиці 5.3 та на рисунку 5.4.')
    add_table(
        doc,
        'Таблиця 5.2 — Результати навантажувального тестування (tests/load-test.js)',
        ['VUs', 'Запитів', 'RPS (зап./с)', 'p50, мс', 'p95, мс', 'p99, мс', 'Помилки, %', 'RSS RAM, МБ'],
        [
            [str(st['virtualUsers']), str(st['totalRequests']), str(st['rps']), str(st['p50Ms']), str(st['p95Ms']), str(st['p99Ms']), str(st['errorRatePercent']), str(st['rssMemoryMb'])]
            for st in load_rep
        ],
        col_widths=[1.6, 2.2, 2.4, 2.0, 2.0, 2.0, 2.2, 2.6]
    )
    add_figure(doc, DIAG_DIR / 'fig_5_1_load_chart.png', 'Рисунок 5.3 — Затримки та пропускна здатність API при навантаженні 5/10/25 VU', width_cm=15.5)
    add_table(
        doc,
        'Таблиця 5.3 — Результати тестування обсягом даних (tests/volume-test.js)',
        ['Обсяг БД (записів)', 'Фільтрований список (avg / p95, мс)', 'Текстовий пошук (avg / p95, мс)', 'Звіт прострочень SLA (avg / p95, мс)'],
        [
            [str(r['recordsCount']), f"{r['listFiltered']['avgMs']} / {r['listFiltered']['p95Ms']}", f"{r['searchByText']['avgMs']} / {r['searchByText']['p95Ms']}", f"{r['overdueReport']['avgMs']} / {r['overdueReport']['p95Ms']}"]
            for r in vol_rep
        ],
        col_widths=[3.5, 4.5, 4.5, 4.5]
    )
    add_figure(doc, DIAG_DIR / 'fig_5_2_volume_chart.png', 'Рисунок 5.4 — Час виконання запитів до SQLite залежно від обсягу даних', width_cm=14.5)

    add_h2(doc, '5.6 Дефекти, повторна перевірка та висновок про якість')
    add_table(
        doc,
        'Таблиця 5.4 — Журнал виявлених дефектів та підтвердження їхнього усунення',
        ['Код', 'Опис дефекту та вплив', 'Пов’язана вимога', 'Виправлення та повторна перевірка'],
        [
            ['DEF-01', 'path.resolve(":memory:") у src/config.js формував некоректний шлях на Windows для тестів у пам’яті.', 'NFR-RUN-01', 'Додано перевірку rawDbPath === ":memory:"; повторний прогін 7/7 тестів успішний. Статус: Перевірено.'],
            ['DEF-02', 'Вбудована функція LOWER() у SQLite не перетворювала кириличні символи в нижній регістр при пошуку («монтаж»).', 'FR-01', 'Зареєстровано власну SQL-функцію uk_lower у connection.js та додано автотест в integration.test.js. Статус: Перевірено.']
        ],
        col_widths=[1.8, 6.2, 2.6, 6.4]
    )

    # =========================================================================
    # ВИСНОВКИ
    # =========================================================================
    add_h1(doc, 'ВИСНОВКИ')
    add_p(doc, 'У межах курсової роботи виконано повний цикл аналізу, проєктування, програмної реалізації та перевірки якості корпоративного сайту із системою керування зверненнями «ВольтаПро Сервіс» (варіант 5). Поставленої мети досягнуто: створено працездатне рішення, що забезпечує прозоре подання та опрацювання заявок із контролем строків SLA.')
    add_p(doc, 'Основні результати: реалізовано модульну серверну частину на Node.js + Express + SQLite, адаптивний інтерфейс на HTML5/CSS3/JS/Bootstrap 5, ідемпотентну реєстрацію звернень, публічний пошук заявки без витоку внутрішніх приміток, рольовий кабінет менеджера (MANAGER, SENIOR_MANAGER), календарний розрахунок дедлайнів SLA у робочих днях та звіт про прострочені заявки. Проведено комплексне тестування (90,00 % покриття рядків, 800,99 RPS при 25 VU та p95 = 7,37 мс, перевірка до 10 000 записів).')
    add_p(doc, 'Обмеженнями поточної навчальної версії є локальний запуск без реального SMTP/SMS-шлюзу та однопроцесна СКБД SQLite. Напрямами розвитку є підключення черги поштових сповіщень та перенесення сховища на PostgreSQL.')

    # =========================================================================
    # ПЕРЕЛІК ДЖЕРЕЛ
    # =========================================================================
    add_h1(doc, 'ПЕРЕЛІК ДЖЕРЕЛ')
    for src in [
        '1. Арсірій О. О., Червоненко П. П. Методичні вказівки до курсової роботи з дисципліни «Веб-технології та веб-дизайн». Одеса: ОНПУ, 2017. 15 с.',
        '2. Методичні вказівки до виконання курсової роботи з дисципліни «Веб-технології та веб-дизайн». Об’єднана кафедральна редакція. Одеса: НУ «Одеська політехніка», 2026. 92 с.',
        '3. ДСТУ 3008:2015. Інформація та документація. Звіти у сфері науки і техніки. Структура та правила оформлювання. Київ: ДП «УкрНДНЦ», 2016.',
        '4. ДСТУ 8302:2015. Інформація та документація. Бібліографічне посилання. Загальні положення та правила складання. Київ: ДП «УкрНДНЦ», 2016.',
        '5. WHATWG. HTML Living Standard. URL: https://html.spec.whatwg.org/multipage/ (дата звернення: 05.11.2026).',
        '6. Bootstrap Documentation. Version 5.3. URL: https://getbootstrap.com/docs/5.3/ (дата звернення: 05.11.2026).',
        '7. Node.js Documentation. SQLite (node:sqlite) & Test Runner. URL: https://nodejs.org/docs/latest/api/ (дата звернення: 10.11.2026).',
        '8. Express.js Documentation. URL: https://expressjs.com/ (дата звернення: 10.11.2026).',
        '9. IETF. RFC 9110: HTTP Semantics. 2022. URL: https://www.rfc-editor.org/rfc/rfc9110.html (дата звернення: 10.11.2026).',
        '10. IETF. RFC 9457: Problem Details for HTTP APIs. 2023. URL: https://www.rfc-editor.org/rfc/rfc9457.html (дата звернення: 10.11.2026).',
        '11. OWASP Foundation. OWASP Top 10:2025 & Password Storage Cheat Sheet. URL: https://owasp.org/ (дата звернення: 12.11.2026).',
        '12. NIST. SP 800-63B-4: Digital Identity Guidelines. 2025. URL: https://pages.nist.gov/800-63-4/sp800-63b.html (дата звернення: 12.11.2026).',
        '13. W3C. Web Content Accessibility Guidelines (WCAG) 2.2. URL: https://www.w3.org/TR/WCAG22/ (дата звернення: 15.11.2026).',
        '14. Docker. Docker Compose Documentation. URL: https://docs.docker.com/compose/ (дата звернення: 15.11.2026).',
        '15. Object Management Group. Unified Modeling Language, version 2.5.1. URL: https://www.omg.org/spec/UML/2.5.1 (дата звернення: 15.11.2026).'
    ]:
        add_p(doc, src, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=3)

    # =========================================================================
    # ДОДАТКИ А, Б, В, Г
    # =========================================================================
    add_h1(doc, 'ДОДАТОК А. МАТРИЦЯ ВІДПОВІДНОСТІ ВИМОГ І РЕЗУЛЬТАТІВ')
    add_table(
        doc,
        'Таблиця А.1 — Матриця простежуваності вимог, реалізації, тестів і доказів (тег v1.0.0)',
        ['Вимога', 'Модуль реалізації', 'Тест / перевірка', 'Доказ', 'Статус'],
        [
            ['FR-01 Каталог', 'serviceRepository.js, connection.js (uk_lower)', 'TC-01; integration.test.js; e2e', 'Рис. 4.2, e2e-ui-report.json', 'Пройдено'],
            ['FR-02, FR-03 Форма', 'inquiryService.js, validationRules.js', 'TC-02, TC-03; integration.test.js', 'Рис. 4.3, 4.4', 'Пройдено'],
            ['FR-04, FR-05 Пошук', 'inquiryService.js (lookupPublicInquiry)', 'TC-04; integration.test.js', 'Рис. 4.5 (без приміток)', 'Пройдено'],
            ['FR-06, FR-07 Статуси', 'inquiryRules.js, manager.js', 'TC-07; unit.test.js', 'Рис. 4.6, 4.7', 'Пройдено'],
            ['FR-08 Права ролей', 'auth.js, inquiryRules.js', 'TC-05, TC-06, TC-08', 'Негативні тести 401/403', 'Пройдено'],
            ['FR-09, FR-10 SLA', 'calendarRules.js, inquiryService.js', 'TC-09, TC-10; unit/integration', 'Рис. 4.8, volume-report.json', 'Пройдено'],
            ['NFR-UI-01 UI', 'styles.css, Bootstrap 5', 'e2e-screenshots.js (360/768/1280)', 'Рис. 5.1, 5.2', 'Пройдено'],
            ['NFR-PERF-01 Швидкодія', 'SQLite WAL, індекси', 'load-test.js, volume-test.js', 'Рис. 5.3, 5.4, load-report.json', 'Пройдено']
        ],
        col_widths=[3.0, 4.4, 3.8, 3.6, 2.2]
    )

    add_h1(doc, 'ДОДАТОК Б. ПАСПОРТИ ТЕСТІВ ТА ЗВЕДЕНИЙ ПРОТОКОЛ ПРИЙМАННЯ')
    add_table(
        doc,
        'Таблиця Б.1 — Паспорт негативного тесту доступу та приватності (TC-04 / TC-06)',
        ['Поле', 'Заповнено виконавцем'],
        [
            ['Ідентифікатор і назва', 'TC-06 — Заборона доступу менеджера до чужої заявки та приховування внутрішніх приміток'],
            ['Пов’язані вимоги', 'FR-05, FR-08, NFR-SEC-01'],
            ['Передумови та дані', 'Створено заявки manager1 та manager2; у заявці є запис із is_internal = 1.'],
            ['Послідовність дій', '1) Публічний запит lookup; 2) Запит manager2 до заявки manager1; 3) Спроба manager1 відкрити завершену заявку знову.'],
            ['Очікуваний / Фактичний результат', 'У lookup внутрішній текст відсутній; для чужої заявки — 403; для повторного відкриття — 403.'],
            ['Статус і доказ', 'Пройдено (tests/integration.test.js, версія v1.0.0).']
        ],
        col_widths=[4.8, 12.2]
    )
    add_table(
        doc,
        'Таблиця Б.2 — Зведений протокол приймання вебсистеми',
        ['Група перевірок', 'Пройдено', 'Не пройдено', 'Заблоковано', 'Коментар'],
        [
            ['Бізнес-сценарії та правила', '7', '0', '0', 'Модульні та інтеграційні тести (90,00 % coverage)'],
            ['Права та захист даних', '5', '0', '0', 'Перевірено 401/403, scrypt та фільтрацію приміток'],
            ['Інтерфейс і доступність', '10', '0', '0', 'Google Chrome E2E на 360, 768, 1280 px'],
            ['Швидкодія й обсяг даних', '6', '0', '0', 'Етапи 5/10/25 VUs та обсяги 100/1000/10000'],
            ['Запуск і документація', '3', '0', '0', 'migrate → seed → start за README.md']
        ],
        col_widths=[4.5, 2.0, 2.4, 2.4, 5.7]
    )

    add_h1(doc, 'ДОДАТОК В. КОНТРОЛЬНИЙ СПИСОК ПЕРЕД ПОДАННЯМ ТА ЖУРНАЛ ЗМІН')
    add_table(
        doc,
        'Таблиця В.1 — Журнал істотних змін та використання допоміжних засобів / ШІ',
        ['Дата', 'Засіб / зміна і причина', 'Використаний результат та перевірка'],
        [
            ['25.11.2026', 'Допоміжний ШІ-інструмент: підготовка початкового каркаса, діаграм, автотестів та верстки ПЗ', 'Усі модулі, цифри навантаження, покриття 90,00 % і скриншоти отримано з реальних запусків та перевірено вручну.'],
            ['25.11.2026', 'Виправлення DEF-02: додано SQL-функцію uk_lower для пошуку українською кирилицею', 'Перевірено ручним пошуком «монтаж» у браузері та автотестом в integration.test.js.']
        ],
        col_widths=[2.8, 6.8, 7.4]
    )

    add_h1(doc, 'ДОДАТОК Г. СУПРОВІДНА ДОКУМЕНТАЦІЯ ПРОГРАМНОГО ПРОДУКТУ')
    add_h2(doc, 'Г.1 Інструкція користувача')
    add_p(doc, 'Для подання звернення відкрийте http://localhost:3000/, оберіть послугу в каталозі, заповніть форму та збережіть виданий номер (VP-YYYY-NNNN) і 6-символьний код доступу. Перевірити статус можна на /status (демо: VP-2026-0001 / DEMO01). Для роботи менеджера перейдіть на /manager (демо-логіни: senior@voltapro.ua, manager1@voltapro.ua, manager2@voltapro.ua; пароль: VoltaPro#Manager2026!).')
    add_h2(doc, 'Г.2 Інструкція зі збирання та запуску')
    add_code_block(
        doc,
        'Лістинг Г.1 — Команди запуску та тестування',
        """npm install
cp .env.example .env
npm run migrate
npm run seed
npm start
# Тести: npm test | npm run test:volume | npm run test:load"""
    )
    add_h2(doc, 'Г.3 Інструкція адміністратора та Г.4 Короткий опис JSON API')
    add_p(doc, 'Конфігурація задається у .env (PORT, DB_PATH, SESSION_TTL_HOURS). База даних SQLite зберігається у data/voltapro.db. Повна специфікація REST API доступна у docs/openapi.yaml та таблиці 3.5.')

    doc.save(OUT_DOCX)
    print('Збережено оновлений DOCX:', OUT_DOCX)


if __name__ == '__main__':
    build_document()
