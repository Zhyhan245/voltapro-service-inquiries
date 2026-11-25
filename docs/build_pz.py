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
    add_p(doc, 'У пояснювальній записці висвітлено етапи аналізу, проєктування, програмування та комплексної перевірки корпоративного вебпорталу компанії «ВольтаПро Сервіс», оснащеного підсистемою обліку та супроводу клієнтських звернень (Варіант 5). Головним завданням проєкту є заміна розрізнених каналів приймання заявок єдиним вебсередовищем, де замовники можуть самостійно відстежувати хід робіт, а менеджери — контролювати регламентні строки виконання (SLA).')
    add_p(doc, 'Серверну логіку реалізовано на платформі Node.js (v26.13.1) із використанням фреймворку Express 4.21.2. Як сховище обрано вбудовану реляційну СКБД SQLite (модуль node:sqlite у режимі журналювання WAL із контролем зовнішніх ключів, обмеженнями CHECK та атомарними транзакціями BEGIN IMMEDIATE). Клієнтську частину створено засобами HTML5, CSS3, модульного JavaScript (ES6+) і локального пакета Bootstrap 5. Серед ключових можливостей — пошуковий каталог сервісів (із власною функцією uk_lower для коректної обробки української кирилиці), захист від дублювання заявок через ключ ідемпотентності, публічна перевірка статусу без розкриття внутрішніх нотаток персоналу, рольовий кабінет (MANAGER, SENIOR_MANAGER), обчислення дедлайнів у робочих днях за виробничим календарем та формування звіту про прострочені звернення.')
    add_p(doc, 'Надійність продукту доведено повним проходженням 7 груп модульних й інтеграційних тестів (покриття 90,00 % рядків коду), 10 наскрізними сценаріями в браузері Google Chrome на екранах 360, 768 та 1280 пікселів, навантажувальними випробуваннями (800,99 зап./с при 25 віртуальних користувачах, p95 = 7,37 мс, 0 % відмов) та замірами швидкодії на масиві до 10 000 записів.')
    add_p(doc, 'Ключові слова: корпоративний вебпортал, супровід звернень, календарний SLA, Node.js, Express, SQLite, Bootstrap 5, рольовий доступ.')

    add_h1(doc, 'ABSTRACT')
    add_p(doc, 'This coursework report presents the analysis, architecture, implementation, and verification of a corporate web portal with a client inquiry tracking subsystem for the engineering company "VoltaPro Service" (Variant 5). The platform replaces fragmented communication channels with a centralized web workflow that gives customers transparent status lookup and provides staff with automated SLA deadline monitoring.')
    add_p(doc, 'The server side runs on Node.js (v26.13.1) and Express 4.21.2, backed by an embedded SQLite database (node:sqlite in WAL mode with foreign keys, CHECK constraints, and BEGIN IMMEDIATE transactions). The frontend is built with semantic HTML5, CSS3, ES6+ JavaScript modules, and locally bundled Bootstrap 5. Core features cover a searchable service catalog (enhanced with a custom uk_lower SQL function for Cyrillic case-folding), idempotent inquiry submission, public status tracking that strips internal staff notes, role-based access control (MANAGER and SENIOR_MANAGER), business-day SLA calculation against a holiday calendar, and an overdue report.')
    add_p(doc, 'Verification results demonstrate 7/7 passing unit and integration test suites (90.00% line coverage), 10 passing Chrome end-to-end checks at 360, 768, and 1280 px viewports, stable throughput under load (800.99 RPS at 25 VUs, p95 = 7.37 ms, zero errors), and predictable SQLite query scaling up to 10,000 records.')
    add_p(doc, 'Keywords: corporate web portal, inquiry tracking, business-day SLA, Node.js, Express, SQLite, Bootstrap 5, role-based access.')

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
    add_p(doc, 'Швидкість і прозорість обробки запитів від замовників напряму визначають репутацію та операційну ефективність інженерно-сервісних компаній. На практиці звернення часто надходять різними каналами — телефоном, електронною поштою або через месенджери, через що частина деталей губиться, а клієнт залишається в невідомості щодо поточного стану своєї заявки.')
    add_p(doc, 'Типові контактні форми на корпоративних сторінках не забезпечують повноцінного циклу супроводу: вони не генерують унікальних кодів відстеження, не блокують випадкових повторних відправлень, не ізолюють внутрішні нотатки працівників від зовнішнього перегляду та не вміють рахувати нормативні строки виконання з урахуванням вихідних і свят.')
    add_p(doc, 'Метою курсової роботи є підвищення керованості та прозорості клієнтського сервісу підприємства «ВольтаПро Сервіс» шляхом створення вебсистеми, що поєднує каталог послуг, захищену реєстрацію звернень, онлайн-перевірку статусу, рольовий кабінет персоналу та календарний контроль SLA.')
    add_p(doc, 'Для досягнення окресленої мети в ході проєктування розв’язано такі задачі:')
    for t in [
        'проаналізувати бізнес-процеси опрацювання заявок, зіставити три сучасні платформи підтримки та сформувати перелік вимог;',
        'розробити алгоритми підрахунку дедлайнів SLA у робочих днях, валідації життєвого циклу статусів та ідемпотентного збереження;',
        'спроєктувати чотирирівневу серверну структуру, схему реляційної БД SQLite і контракти REST API;',
        'реалізувати програмний комплекс на стеку Node.js, Express, SQLite та Bootstrap 5 із розмежуванням прав і звітом прострочень;',
        'провести модульне, інтеграційне, браузерне, навантажувальне та об’ємне тестування й оформити звітні матеріали.'
    ]:
        add_bullet(doc, t)
    add_p(doc, 'Об’єктом розробки виступає процес взаємодії клієнта та сервісного підрозділу на корпоративному сайті. Предмет роботи складають структури даних, доменні правила переходів статусів, механізм виробничого календаря SLA та інтерфейси обміну між браузером і сервером.')
    add_p(doc, 'Навчальний характер проєкту передбачає використання синтетичних даних компанії «ВольтаПро Сервіс», демонстраційних облікових записів (senior@voltapro.ua, manager1@voltapro.ua, manager2@voltapro.ua) та тестових звернень без підключення реальних шлюзів розсилки SMS чи електронної пошти.')

    # =========================================================================
    # РОЗДІЛ 1. АНАЛІЗ ПРЕДМЕТНОЇ ОБЛАСТІ ТА ВИМОГИ
    # =========================================================================
    add_h1(doc, '1 АНАЛІЗ ПРЕДМЕТНОЇ ОБЛАСТІ ТА ФОРМУВАННЯ ВИМОГ')

    add_h2(doc, '1.1 Опис проблемної ситуації, аудиторії та огляд аналогів')
    add_p(doc, 'За відсутності спеціалізованої вебсистеми звернення до сервісної компанії «ВольтаПро Сервіс» фіксувалися розрізнено — телефоном і через загальну електронну пошту. Такий підхід породжує низку операційних ризиків: клієнт не отримує реєстраційного номера для контролю виконання, часто надсилає повторні запити, працівники можуть ненароком оприлюднити внутрішні службові коментарі, а дотримання регламентних строків відстежується вручну без урахування святкових днів.')
    add_p(doc, 'Коло потенційних користувачів поділяється на дві категорії: зовнішні відвідувачі (фізичні особи та корпоративні клієнти), зацікавлені в ознайомленні з переліком послуг, швидкій подачі заявки та перевірці її статусу, а також внутрішній персонал (менеджери й старший менеджер), який розподіляє навантаження, веде робочу історію та моніторить прострочення SLA.')
    add_p(doc, 'Щоб виявити вдалі галузеві рішення, проаналізовано документацію трьох поширених платформ обслуговування заявок (станом на 05.11.2026): Zendesk Support, Bitrix24 CRM і Jira Service Management. Порівняння зведено в таблиці 1.1.')

    add_table(
        doc,
        'Таблиця 1.1 — Зіставлення функціональних можливостей існуючих платформ та цільової системи',
        ['Аспект порівняння', 'Zendesk Support', 'Bitrix24 CRM', 'Jira Service Management', 'Рішення для «ВольтаПро Сервіс»'],
        [
            ['Орієнтація на користувача', 'Великі й середні центри підтримки', 'Комерційні відділи та малий бізнес', 'Сервісні підрозділи та ІТ-служби (ITSM)', 'Сфокусована трійка ролей: клієнт, менеджер, старший менеджер'],
            ['Організація потоку заявок', 'Черги тикетів із правилами маршрутизації', 'Ліди з вебформ усередині CRM-воронки', 'Портал запитів із маршрутами погодження', 'Єдиний потік: каталог → форма звернення → пошук статусу → кабінет + SLA'],
            ['Приватність і обробка відмов', 'Чітке відділення публічних відповідей від internal notes', 'Деталі приховані в закритих платних модулях', 'Обов’язкові поля при переходах станів workflow', 'Помилки RFC 9457, збереження введених даних, жорстке фільтрування нотаток'],
            ['Мобільність і доступність', 'Адаптивний вебінтерфейс із підтримкою WCAG', 'Вебпортал і окремі мобільні додатки', 'Адаптивний клієнтський портал', 'Коректна верстка на 360, 768, 1280 px і повне керування з клавіатури'],
            ['Причини обмеженої придатності', 'Висока вартість ліцензій і складне налаштування', 'Закрита хмарна інфраструктура та підписка', 'Складний інтерфейс для звичайних замовників', 'Автономний локальний стек, відкритий код, прозорий виробничий календар']
        ],
        col_widths=[2.8, 3.5, 3.5, 3.5, 3.7]
    )

    add_p(doc, 'Проведене порівняння показує: комерційні платформи вимагають значних витрат і надлишкових налаштувань, однак формують корисні стандарти — суворе відокремлення службових нотаток (Zendesk), примусове зазначення причини при відхиленні (Jira) та прив’язку строків до робочого календаря. Саме ці підходи закладено в основу власного застосунку.')

    add_h2(doc, '1.2 Рольова модель, повноваження та межі розробки')
    add_p(doc, 'Архітектура доступу спирається на три ролі (таблиця 1.2): зовнішній Гість / Клієнт, Менеджер (MANAGER) і Старший менеджер (SENIOR_MANAGER). Перевірка прав виконується виключно на боці сервера перед виконанням будь-якої чутливої дії.')

    add_table(
        doc,
        'Таблиця 1.2 — Розподіл прав доступу між ролями користувачів',
        ['Функція / API-маршрут', 'Гість / Клієнт', 'Менеджер (MANAGER)', 'Старший менеджер (SENIOR_MANAGER)', 'Спосіб серверного контролю'],
        [
            ['Перегляд каталогу (GET /api/services)', 'Дозволено', 'Дозволено', 'Дозволено', 'Відкритий маршрут; фільтр is_active = 1'],
            ['Реєстрація заявки (POST /api/inquiries)', 'Дозволено', 'Дозволено', 'Дозволено', 'Валідація DTO, idempotency_key, обчислення due_date'],
            ['Перевірка стану (GET /api/inquiries/lookup)', 'За кодом', 'За кодом', 'За кодом', 'Звірка ticketNumber + accessCode; вибірка is_internal = 0'],
            ['Перелік заявок (GET /api/manager/inquiries)', 'Відмова 401', 'Свої та вільні заявки', 'Повний перелік', 'requireAuth + фільтрація за assigned_employee_id'],
            ['Деталі заявки (GET /api/manager/inquiries/:id)', 'Відмова 401', 'Доступні (чужі → 403)', 'Усі без винятку', 'Перевірка у assertEmployeeCanAccessInquiry()'],
            ['Призначення виконавця (PATCH .../assign)', 'Відмова 401', 'Тільки на себе (403 іншим)', 'На довільного або зняття', 'Перевірка ролі всередині assignInquiry()'],
            ['Зміна робочих статусів', 'Відмова 401', 'Для доступних заявок', 'Для кожної заявки', 'validateStatusTransition() + контроль причини відмови'],
            ['Повернення закритої заявки в роботу', 'Відмова 401', 'Заборонено (403)', 'Дозволено', 'Окрема перевірка ролі SENIOR_MANAGER для фінальних станів'],
            ['Аналітика SLA (GET .../reports/overdue)', 'Відмова 401', 'За своїми та вільними', 'Зведений звіт', 'requireAuth + розрахунок evaluateSlaState()']
        ],
        col_widths=[4.4, 2.3, 3.3, 3.3, 3.7]
    )

    add_p(doc, 'До меж реалізації включено публічний вебсайт, адміністративний кабінет, REST API, доменне ядро, базу даних SQLite і конфігурацію контейнеризації. Інтерфейсна бібліотека Bootstrap 5 розгортається локально без звернення до зовнішніх CDN, а інтеграція з реальними поштовими чи SMS-провайдерами не передбачається.')

    add_h2(doc, '1.3 Специфікація функціональних вимог')
    add_p(doc, 'Усі вимоги до поведінки застосунку зведено в таблиці 1.3 за шістьма напрямами (каталог послуг, приймання звернень, публічний пошук, робочий процес статусів, ізоляція доступу та календарний SLA). Детальні паспорти двох ключових функцій наведено в таблицях 1.4 і 1.5.')

    add_table(
        doc,
        'Таблиця 1.3 — Перелік функціональних вимог до програмної системи',
        ['Ідентифікатор', 'Група', 'Зміст вимоги та умови приймання'],
        [
            ['FR-01', 'F-01 Каталог', 'Текстовий пошук (з підтримкою регістронезалежної кирилиці) та фільтр категорій показують лише активні послуги.'],
            ['FR-02', 'F-02 Звернення', 'Помилкові або неповні контакти відхиляються зі статусом 400 VALIDATION_ERROR (fieldErrors), не очищаючи введені поля форми.'],
            ['FR-03', 'F-02 Ідемпотентність', 'Повторний запит із ідентичним idempotencyKey не створює нового рядка в БД, а повертає 200 (duplicate: true).'],
            ['FR-04', 'F-03 Пошук', 'Пошук за номером VP-YYYY-NNNN і 6-значним кодом доступу; за невірного коду повертається 404 без витоку відомостей.'],
            ['FR-05', 'F-05 Приватність', 'У публічній відповіді віддається лише зовнішня історія (is_internal = 0); внутрішні примітки персоналу відсіюються.'],
            ['FR-06', 'F-04 Списки', 'Кабінет менеджера підтримує фільтрацію, пошук, посторінковий перегляд і призначення відповідального згідно з роллю.'],
            ['FR-07', 'F-04 Статуси', 'Граф переходів «NEW → IN_PROGRESS → COMPLETED / REJECTED»; відхилення без тексту причини (≥ 5 симв.) блокується (400).'],
            ['FR-08', 'F-05 Доступ', 'Менеджер отримує 403 при спробі відкрити чужу заявку або самостійно поновити вже завершене/відхилене звернення.'],
            ['FR-09', 'F-06 Календар', 'Під час реєстрації заявки дата due_date обчислюється в робочих днях із пропуском вихідних і записів calendar_exceptions.'],
            ['FR-10', 'F-06 Моніторинг', 'Звіт прострочень розраховує кількість затриманих робочих днів для активних заявок і відображає календар винятків.']
        ],
        col_widths=[1.8, 3.8, 11.4]
    )

    add_table(
        doc,
        'Таблиця 1.4 — Паспорт вимоги FR-02 / FR-03 (подання та ідемпотентність звернення)',
        ['Атрибут', 'Опис'],
        [
            ['Код вимоги', 'FR-02 / FR-03'],
            ['Назва', 'Приймання клієнтської заявки з уникненням дублювання'],
            ['Виконавець', 'Гість або зареєстрований клієнт'],
            ['Початкові умови', 'Обрана послуга наявна в таблиці services і має ознаку is_active = 1'],
            ['Вхідні параметри', 'serviceId, contactName, companyName, phone, email, subject, message, idempotencyKey'],
            ['Успішний сценарій', 'Додаються контакт, звернення (NEW) та запис історії; клієнт отримує 201 Created, ticketNumber, accessCode і dueDate'],
            ['Дублювання запиту', 'За наявності того ж idempotencyKey повертається 200 OK (duplicate: true) із реквізитами вже створеної заявки'],
            ['Помилки вводу', 'Статус 400 VALIDATION_ERROR (RFC 9457) із деталізацією fieldErrors; транзакція не фіксується, форма зберігає дані'],
            ['Підтвердження', 'Інтеграційний тест tests/integration.test.js та браузерний прогін у Chrome (рис. 4.3, 4.4)']
        ],
        col_widths=[4.2, 12.8]
    )

    add_table(
        doc,
        'Таблиця 1.5 — Паспорт вимоги FR-07 / FR-08 (керування статусами й авторизація)',
        ['Атрибут', 'Опис'],
        [
            ['Код вимоги', 'FR-07 / FR-08'],
            ['Назва', 'Контроль переходів станів заявки та рольових обмежень'],
            ['Виконавець', 'Працівник із роллю MANAGER або SENIOR_MANAGER'],
            ['Початкові умови', 'Наявна активна сесія; рядовий менеджер може змінювати лише власну або нерозподілену заявку'],
            ['Вхідні параметри', 'inquiryId, status (IN_PROGRESS | COMPLETED | REJECTED), rejectionReason, publicResponse'],
            ['Успішний сценарій', 'Атомарне оновлення статусу, фіксація closed_at для фінальних станів і додавання події до inquiry_history'],
            ['Граничні відмови', 'Відхилення без причини → 400; повторне відкриття без SENIOR_MANAGER → 403; заборонений перехід → 409'],
            ['Підтвердження', 'Набори тестів tests/unit.test.js та tests/integration.test.js']
        ],
        col_widths=[4.2, 12.8]
    )

    add_h2(doc, '1.4 Нефункціональні обмеження та матриця трасування')
    add_p(doc, 'Якісні та експлуатаційні характеристики вебсистеми закріплено в таблиці 1.6. Повна прив’язка кожної вимоги до модулів коду й тестів наведена в Додатку А.')
    add_table(
        doc,
        'Таблиця 1.6 — Нефункціональні вимоги до системи «ВольтаПро Сервіс»',
        ['Код', 'Напрям', 'Критерій якості', 'Метод контролю'],
        [
            ['NFR-SEC-01', 'Інформаційна безпека', 'Хешування паролів scrypt із сіллю (NIST SP 800-63B-4, ≥ 15 симв.); cookie з HttpOnly і SameSite=Lax; захист від IDOR (403) і SQL-ін’єкцій.', 'unit.test.js, integration.test.js, аналіз коду.'],
            ['NFR-DATA-01', 'Надійність збереження', 'Відтворення структури міграцією 001_init.sql; збереження даних після перезапуску; активні зовнішні ключі та CHECK-правила SQLite.', 'migrate.js, seed.js, контрольний рестарт.'],
            ['NFR-UI-01', 'Зручність і WCAG 2.2', 'Коректна ширина без горизонтальної прокрутки на 360, 768, 1280 px; помітний фокус (:focus-visible), skip-link, текстові мітки станів.', 'e2e-screenshots.js у Chrome + візуальний огляд.'],
            ['NFR-PERF-01', 'Швидкодія', 'При навантаженні 5, 10, 25 VU показник p95 затримки API не перевищує 800 мс, частка помилок < 1 %.', 'load-test.js та volume-test.js (до 10 000 рядків).'],
            ['NFR-RUN-01', 'Відтворюваність', 'Повний запуск за кроками README.md у чистому середовищі без зашитих секретів (або через Docker Compose).', 'Чиста перевірка npm install → migrate → seed → start.']
        ],
        col_widths=[2.5, 2.8, 7.2, 4.5]
    )

    add_h2(doc, '1.5 Діаграма прецедентів')
    add_p(doc, 'Функціональні сценарії трьох ролей користувачів узагальнено на UML-діаграмі прецедентів [15] (рисунок 1.1).')
    add_figure(doc, DIAG_DIR / 'fig_1_1_use_cases.png', 'Рисунок 1.1 — UML-діаграма прецедентів вебсистеми «ВольтаПро Сервіс»', width_cm=15.0)
    add_p(doc, 'Ключовий прецедент «Подання звернення» ініціюється гостем для активної послуги: у разі валідних даних створюються контакт і запис зі статусом NEW, розраховується дедлайн SLA в робочих днях та видаються номер і код доступу; при помилках повертаються деталі валідації без втрати тексту форми, а повторний запит із тим самим ключем ідемпотентності повертає вже наявне звернення.')

    add_h2(doc, '1.6 Верифікація якості вимог')
    add_table(
        doc,
        'Таблиця 1.7 — Оцінка сформульованих вимог за критеріями якості',
        ['Критерій', 'Зміст перевірки', 'Обґрунтування виконання у проєкті'],
        [
            ['Атомарність', 'Кожен пункт фіксує єдине правило', 'Окремо винесені валідація (FR-02), ідемпотентність (FR-03), фільтрація нотаток (FR-05), статуси (FR-07) і SLA (FR-09).'],
            ['Тестованість', 'Наявні чіткі вхідні та вихідні умови', 'Для кожної вимоги визначено HTTP-код, тип помилки RFC 9457 та перевіряючий автотест.'],
            ['Однозначність', 'Узгоджений словник станів і ролей', 'Зафіксовано статуси NEW, IN_PROGRESS, COMPLETED, REJECTED та ролі MANAGER, SENIOR_MANAGER.'],
            ['Повнота', 'Враховано штатні й аварійні гілки', 'Охоплено помилки контакту, повторне надсилання, невірний код пошуку, чужі заявки, відмову без пояснення.'],
            ['Узгодженість', 'Відсутність взаємних суперечностей', 'Правила життєвого циклу й доступу зосереджено в єдиному модулі src/domain/inquiryRules.js.'],
            ['Трасованість', 'Наскрізний зв’язок із кодом і доказами', 'Задокументовано в таблиці А.1 (Додаток А).']
        ],
        col_widths=[3.2, 5.2, 8.6]
    )

    # =========================================================================
    # РОЗДІЛ 2. АЛГОРИТМІЧНЕ ЗАБЕЗПЕЧЕННЯ ВЕБСИСТЕМИ
    # =========================================================================
    add_h1(doc, '2 АЛГОРИТМІЧНЕ ТА МАТЕМАТИЧНЕ ЗАБЕЗПЕЧЕННЯ')

    add_h2(doc, '2.1 Опис базових алгоритмів і їхні характеристики')
    add_p(doc, 'Роботу доменної логіки забезпечують чотири провідні алгоритми: обчислення планової дати виконання SLA за робочим календарем (addBusinessDays); визначення поточного стану SLA та тривалості затримки (evaluateSlaState); валідація допустимості переходу між статусами з урахуванням ролі (validateStatusTransition); та ідемпотентна реєстрація нового звернення в межах транзакції (createPublicInquiry). Їхні специфікації узагальнено в таблиці 2.1.')

    add_table(
        doc,
        'Таблиця 2.1 — Характеристики провідних алгоритмів програмного комплексу',
        ['Назва алгоритму', 'Пов’язані вимоги', 'Вхідні й вихідні дані', 'Ключові інваріанти та обробка винятків', 'Оцінка складності'],
        [
            ['1. addBusinessDays', 'FR-09 (календар SLA)', 'Вхід: startDate, slaDays (1..30), exceptionsMap. Вихід: dueDate (YYYY-MM-DD).', 'Результуюча дата завжди припадає на робочий день; інтервал містить рівно slaDays робочих діб. Запобіжник guard = 3660.', 'Час: O(d), d ≤ 30; пам’ять: O(1).'],
            ['2. evaluateSlaState', 'FR-10 (контроль строків)', 'Вхід: inquiry, referenceDate, exceptionsMap. Вихід: {state, overdueBusinessDays, remainingBusinessDays}.', 'Класифікує ON_TIME, DUE_TODAY, OVERDUE, CLOSED_ON_TIME, CLOSED_OVERDUE з ігноруванням вихідних і свят.', 'Час: O(Δt) діб; пам’ять: O(1).'],
            ['3. validateStatusTransition', 'FR-07, FR-08 (статуси)', 'Вхід: currentStatus, nextStatus, employeeRole, rejectionReason. Вихід: reason або DomainError.', 'Блокує некоректні переходи (409), поновлення без ролі SENIOR_MANAGER (403) і відмову без пояснення (400).', 'Час: O(1); пам’ять: O(1).'],
            ['4. createPublicInquiry', 'FR-02, FR-03 (створення)', 'Вхід: body (DTO). Вихід: {duplicate, ticketNumber, accessCode, dueDate}.', 'Виконується у BEGIN IMMEDIATE; індекс по idempotency_key унеможливлює появу дублікатів.', 'Час: O(log n + d); пам’ять: O(1).']
        ],
        col_widths=[3.2, 3.2, 3.8, 4.3, 2.5]
    )

    add_p(doc, 'Для активного звернення за умови перевищення звітною датою t_ref планового строку t_due кількість прострочених робочих днів обчислюється за таким виразом:')
    add_p(doc, 'K_overdue(t_due, t_ref) = Σ I_work(t_i),   де t_i ∈ (t_due; t_ref]', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=4)
    add_p(doc, 'Тут t_due позначає розрахований дедлайн (YYYY-MM-DD); t_ref — дату звітування; I_work(t_i) ∈ {0, 1} — індикатор робочої доби, що набуває значення 1 для робочих днів (із пріоритетним урахуванням винятків у calendar_exceptions) і 0 для субот, неділь та державних свят.')

    add_h2(doc, '2.2 Формалізація псевдокодом і блок-схемами')
    add_code_block(
        doc,
        'Лістинг 2.1 — Псевдокод обчислення дедлайну addBusinessDays (src/domain/calendarRules.js)',
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

    add_h2(doc, '2.3 Теоретична складність і практичні заміри')
    add_p(doc, 'Оскільки норматив SLA обмежено діапазоном d ≤ 30 робочих днів, цикл у функції addBusinessDays виконує не більше приблизно 45 кроків навіть за наявності святкових перенесень, тобто працює за сталий час O(1) і не потребує додаткової пам’яті. У формуванні звіту прострочень (getOverdueReport) для n відкритих заявок вибірка спирається на складений індекс idx_inquiries_status_due(status, due_date), оцінка SLA кожного запису виконується за O(1) у межах горизонту, а подальше впорядкування за спаданням затримки через TimSort у рушії V8 має асимптотику O(n log n) за часом і O(n) за пам’яттю.')
    add_p(doc, 'Практичні випробування на стенді (докладно в підрозділі 5.5) засвідчили високу стійкість до зростання обсягів: при збільшенні таблиці у 100 разів (зі 100 до 10 000 рядків) середній час побудови фільтрованого списку змінився з 0,295 мс до 11,297 мс, а звіту прострочених заявок — із 0,527 мс до 53,691 мс.')

    # =========================================================================
    # РОЗДІЛ 3. ПРОЄКТУВАННЯ ВЕБСИСТЕМИ
    # =========================================================================
    add_h1(doc, '3 АРХІТЕКТУРНЕ ПРОЄКТУВАННЯ ВЕБСИСТЕМИ')

    add_h2(doc, '3.1 Архітектурний стиль і технологічні обмеження')
    add_p(doc, 'Програмний комплекс «ВольтаПро Сервіс» спроєктовано як структурований клієнт-серверний моноліт: інтерфейс у браузері (HTML5, CSS3, ES6+ та Bootstrap 5) взаємодіє через JSON REST API із серверним процесом Node.js + Express, який монопольно керує базою SQLite. Такий вибір забезпечує надійні локальні транзакції в межах єдиної БД без накладних витрат на розподілену координацію мікросервісів.')
    add_table(
        doc,
        'Таблиця 3.1 — Архітектурні обмеження та способи їхнього дотримання',
        ['Обмеження', 'Підстава', 'Реалізація і контроль'],
        [
            ['AC-01: Нативний фронтенд (HTML5/CSS3/ES6/Bootstrap 5)', 'Відмова від важких SPA-фреймворків', 'Модулі ES6 у каталозі public/js/, локальне підключення Bootstrap 5 без зовнішніх CDN.'],
            ['AC-02: Node.js + Express + вбудований node:sqlite', 'Обраний серверний стек', 'Відсутність зовнішніх нативних збірок; перевірено в середовищі Node.js v26.13.1.'],
            ['AC-03: Автономність без платних хмарних шлюзів', 'Навчальні межі роботи', 'Усі операції виконуються локально; аудит дій фіксується в журналі БД.'],
            ['AC-04: Готовність до контейнеризації Docker Compose', 'Вимога відтворюваності середовища', 'Наявні Dockerfile та docker-compose.yml із підключенням постійного тому ./data.']
        ],
        col_widths=[5.5, 4.5, 7.0]
    )

    add_h2(doc, '3.2 Структура інтерфейсу користувача (UI View)')
    add_p(doc, 'Графічний інтерфейс поділено на три маршрути (/, /status, /manager), що охоплюють вісім основних станів (рисунок 3.1, таблиця 3.2). Кожен екран підтримує індикацію завантаження, інформативну заглушку при порожніх результатах, підсвічування помилок зі збереженням введеного тексту та реакцію на відмови доступу 401/403 на контрольних ширинах 360, 768 і 1280 пікселів.')
    add_figure(doc, DIAG_DIR / 'fig_3_1_ui_state_map.png', 'Рисунок 3.1 — Карта екранів та станів користувацького інтерфейсу (UI View)', width_cm=15.5)
    add_table(
        doc,
        'Таблиця 3.2 — Характеристика екранів і станів вебінтерфейсу',
        ['Екран', 'Ролі', 'Ключові сценарії та елементи керування'],
        [
            ['UI-01: Вітрина послуг (/)', 'Гість, Персонал', 'GET /api/services; пошуковий рядок, селектор категорій, сітка карток послуг, стан порожнього пошуку.'],
            ['UI-02 / UI-03: Подання та результат (/)', 'Гість, Персонал', 'POST /api/inquiries; валідація полів, захист від подвійного кліку, збереження вводу при 400, картка номера й коду.'],
            ['UI-04: Трекінг заявки (/status)', 'Гість / Клієнт', 'GET /api/inquiries/lookup; поточний статус, дедлайн SLA, публічна стрічка без службових приміток.'],
            ['UI-05: Автентифікація (/manager)', 'Персонал', 'POST /api/auth/login; форма входу, обробка 401, панель підказки навчальних обліковок.'],
            ['UI-06 / UI-07 / UI-08: Робоче місце та SLA (/manager)', 'MANAGER, SENIOR_MANAGER', 'Таблиця з фільтрами й пагінацією, модальна картка з історією та примітками, окрема вкладка звіту SLA.']
        ],
        col_widths=[4.6, 3.4, 9.0]
    )

    add_h2(doc, '3.3 Логічна структура системи (Logical View)')
    add_p(doc, 'Логічна схема (рисунок 3.2) розмежовує клієнтське середовище браузера та довірений серверний контур Node.js + Express, що ізолює файл бази даних SQLite від прямого доступу.')
    add_figure(doc, DIAG_DIR / 'fig_3_2_logical_view.png', 'Рисунок 3.2 — Структурна схема вебсистеми та її довіреної частини (Logical View)', width_cm=15.0)

    add_h2(doc, '3.4 Топологія розгортання (Deployment View)')
    add_p(doc, 'Схема розгортання (рисунок 3.3, таблиця 3.3) ілюструє зв’язок клієнтського браузера на ОС Windows 11 Pro із локальним сервером Node.js (порт 3000) та базою SQLite у файлі data/voltapro.db, а також альтернативний запуск через Docker Compose.')
    add_figure(doc, DIAG_DIR / 'fig_3_3_deployment_view.png', 'Рисунок 3.3 — Схема розгортання вебсистеми «ВольтаПро Сервіс» (Deployment View)', width_cm=14.5)
    add_table(
        doc,
        'Таблиця 3.3 — Параметри вузлів середовища виконання',
        ['Вузол', 'Технічна конфігурація', 'Підтвердження працездатності'],
        [
            ['Клієнтський браузер', 'Windows 11 Pro, Google Chrome, перегляд на 360/768/1280 px за адресою http://localhost:3000', 'e2e-screenshots.js (10/10 успішних перевірок)'],
            ['Сервер застосунку', 'Node.js v26.13.1, Express 4.21.2, порт 3000, змінні середовища з .env', 'npm start + GET /api/health → 200 OK'],
            ['База даних', 'SQLite (node:sqlite) у режимі WAL, файл data/voltapro.db, міграція 001_init.sql', 'npm run migrate + повторний npm run seed']
        ],
        col_widths=[2.8, 7.6, 6.6]
    )

    add_h2(doc, '3.5 Багатошарова декомпозиція (Design View)')
    add_p(doc, 'Код бекенду організовано у чотири ізольовані шари (рисунок 3.4): транспортний рівень (routes, middleware), рівень прикладних сценаріїв (services), чисте доменне ядро (domain) та адаптери доступу до даних й інфраструктури (repositories, db, logger).')
    add_figure(doc, DIAG_DIR / 'fig_3_4_design_layers.png', 'Рисунок 3.4 — Розподіл відповідальності між логічними шарами (Design View)', width_cm=14.5)

    add_h2(doc, '3.6 Динаміка обробки запитів (Process View)')
    add_p(doc, 'Діаграма послідовності (рисунок 3.5) деталізує сценарій подання звернення з урахуванням трьох можливих розгалужень: повернення існуючого запису за ключем ідемпотентності, транзакційне збереження через BEGIN IMMEDIATE та відхилення помилкових даних зі статусом 400 Bad Request.')
    add_figure(doc, DIAG_DIR / 'fig_3_5_sequence_create.png', 'Рисунок 3.5 — Діаграма послідовності реєстрації звернення (Process View)', width_cm=15.5)

    add_h2(doc, '3.7 Інформаційна модель та схема БД (Data View)')
    add_p(doc, 'Реляційна структура SQLite охоплює 7 взаємопов’язаних таблиць (рисунок 3.6): services, client_contacts, inquiries, inquiry_history, employees, sessions і calendar_exceptions. Транспортні структури DTO відокремлено від таблиць зберігання (рисунок 3.7), а атрибути центральної таблиці inquiries описано в таблиці 3.4.')
    add_figure(doc, DIAG_DIR / 'fig_3_6_er_model.png', 'Рисунок 3.6 — ER-модель реляційної бази даних SQLite (Data View)', width_cm=15.5)
    add_figure(doc, DIAG_DIR / 'fig_3_7_domain_dto.png', 'Рисунок 3.7 — Діаграма основних доменних об’єктів та DTO', width_cm=15.5)
    add_table(
        doc,
        'Таблиця 3.4 — Структура полів таблиці звернень inquiries',
        ['Атрибут', 'Тип даних і обмеження', 'Семантика поля'],
        [
            ['id', 'INTEGER PRIMARY KEY AUTOINCREMENT', 'Первинний сурогатний ключ запису'],
            ['ticket_number', 'TEXT NOT NULL UNIQUE', 'Зовнішній номер заявки формату VP-YYYY-NNNN'],
            ['access_code', 'TEXT NOT NULL', '6-символьний секретний код для клієнтського пошуку'],
            ['service_id / contact_id', 'INTEGER NOT NULL FK', 'Посилання на обрану послугу та контактну особу'],
            ['assigned_employee_id', 'INTEGER FK (NULLABLE)', 'Призначений менеджер-виконавець'],
            ['status', 'TEXT CHECK IN (...)', 'Поточний стан: NEW, IN_PROGRESS, COMPLETED, REJECTED'],
            ['rejection_reason', 'TEXT (NULLABLE)', 'Пояснення причини відхилення заявки'],
            ['idempotency_key', 'TEXT UNIQUE', 'Унікальний маркер запобігання дублікатам'],
            ['sla_business_days / due_date', 'INTEGER / TEXT', 'Норматив SLA у робочих днях та обчислений дедлайн'],
            ['created_at / updated_at / closed_at', 'TEXT (ISO 8601)', 'Фіксація часу створення, оновлення та закриття']
        ],
        col_widths=[4.8, 5.2, 7.0]
    )

    add_h2(doc, '3.8 Контракти REST API (Interface View)')
    add_p(doc, 'Специфікацію програмних інтерфейсів задокументовано у форматі OpenAPI 3.1 (файл docs/openapi.yaml). Перелік ключових кінцевих точок подано в таблиці 3.5; повідомлення про помилки формуються за стандартом RFC 9457 Problem Details.')
    add_table(
        doc,
        'Таблиця 3.5 — Основні маршрути та контракти REST API',
        ['HTTP-метод і шлях', 'Рівень доступу', 'Операція та коди відповідей'],
        [
            ['GET /api/health', 'Відкритий', 'Діагностика стану сервера й доступності SQLite (200 / 503)'],
            ['GET /api/services', 'Відкритий', 'Отримання каталогу послуг із фільтрацією й пошуком (200)'],
            ['POST /api/inquiries', 'Відкритий', 'Створення заявки з підтримкою ідемпотентності (201 / 200 / 400 / 404)'],
            ['GET /api/inquiries/lookup', 'За кодом доступу', 'Публічний перегляд стану без службових приміток (200 / 400 / 404)'],
            ['POST /api/auth/login | logout', 'Персонал', 'Автентифікація scrypt, видача/видалення HttpOnly cookie (200 / 401 | 204)'],
            ['GET /api/manager/inquiries', 'MANAGER / SENIOR', 'Список звернень із фільтрами й пагінацією (200 / 401)'],
            ['PATCH /api/manager/inquiries/:id/assign', 'MANAGER / SENIOR', 'Призначення виконавця з урахуванням ролі (200 / 403)'],
            ['PATCH /api/manager/inquiries/:id/status', 'MANAGER / SENIOR', 'Переведення статусу заявки (200 / 400 / 403 / 409)'],
            ['POST /api/manager/inquiries/:id/comments', 'MANAGER / SENIOR', 'Додавання публічної відповіді або внутрішньої нотатки (201 / 400 / 403)'],
            ['GET /api/manager/reports/overdue', 'MANAGER / SENIOR', 'Аналітичний звіт прострочених заявок і святкових днів (200 / 401)']
        ],
        col_widths=[6.0, 3.6, 7.4]
    )

    add_h2(doc, '3.9 Механізми захисту інформації (Security View)')
    add_p(doc, 'Заходи безпеки узгоджено з рекомендаціями OWASP Top 10:2025: для паролів використано алгоритм scrypt (N=16384, r=8, p=1) із 16-байтною сіллю, порівнянням через timingSafeEqual і мінімальною довжиною 15 символів (NIST SP 800-63B-4); керування сесіями спирається на серверну таблицю sessions із HttpOnly та SameSite=Lax cookie; перевірка повноважень на рівні кожного запису виконується в assertEmployeeCanAccessInquiry; захист від ін’єкцій гарантовано параметризованими SQL-запитами, функцією escapeHtml у браузері та HTTP-заголовками Helmet.')

    add_h2(doc, '3.10 Експлуатаційний контур і тестовий стенд')
    add_p(doc, 'Спостережуваність системи забезпечується структурованим JSON-логуванням із наскрізним requestId (без запису секретів) і діагностичною точкою /api/health, що виконує тестовий запит SELECT 1 до БД. Експериментальні вимірювання проводилися на робочій станції під керуванням Microsoft Windows 11 Pro (процесор AMD Ryzen 7 9700X, 8 ядер / 16 потоків, 61,61 ГБ ОЗП, Node.js v26.13.1, npm 11.20.0, Google Chrome).')

    add_h2(doc, '3.11 Постачання та розгортання коду')
    add_p(doc, 'Вихідні тексти та конфігурація супроводжуються в репозиторії GitHub (https://github.com/Zhyhan245/voltapro-service-inquiries, релізний тег v1.0.0) разом із файлами README.md, .env.example, скриптами міграції, початкового наповнення, тестування та Docker Compose.')

    # =========================================================================
    # РОЗДІЛ 4. РЕАЛІЗАЦІЯ ВЕБСИСТЕМИ
    # =========================================================================
    add_h1(doc, '4 ПРОГРАМНА РЕАЛІЗАЦІЯ ТА МЕТРИКИ КОДУ')

    add_h2(doc, '4.1 Модульна організація бекенду')
    add_p(doc, 'Взаємозв’язки між програмними компонентами серверної частини проілюстровано на рисунку 4.1, а фрагмент транзакційної логіки збереження звернення наведено в лістингу 4.1.')
    add_figure(doc, DIAG_DIR / 'fig_4_1_architecture_classes.png', 'Рисунок 4.1 — Ключові модулі та контракти серверної реалізації', width_cm=14.0)
    add_code_block(
        doc,
        'Лістинг 4.1 — Атомарне створення контакту, звернення та історії (src/repositories/inquiryRepository.js)',
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

    add_h2(doc, '4.2 Візуалізація інтерфейсу в робочих сценаріях')
    add_p(doc, 'Знімки екрана працюючого застосунку в Google Chrome (рисунки 4.2–4.8) демонструють каталог послуг, збереження полів форми при валідаційній помилці, видачу коду доступу, клієнтський пошук без службових нотаток, список заявок у кабінеті, модальне вікно обробки та звіт прострочень SLA.')
    add_figure(doc, SCR_DIR / '01-public-desktop-1280.png', 'Рисунок 4.2 — Публічний сайт, каталог послуг та форма звернення (1280 px)', width_cm=14.5)
    add_figure(doc, SCR_DIR / '03-public-form-error-validation.png', 'Рисунок 4.3 — Помилка валідації при неповному контакті зі збереженням полів', width_cm=14.5)
    add_figure(doc, SCR_DIR / '04-public-inquiry-confirmed.png', 'Рисунок 4.4 — Картка підтвердження прийняття звернення з кодом доступу та дедлайном SLA', width_cm=14.5)
    add_figure(doc, SCR_DIR / '05-status-lookup-public.png', 'Рисунок 4.5 — Сторінка публічної перевірки статусу без витоку внутрішніх приміток', width_cm=14.5)
    add_figure(doc, SCR_DIR / '07-manager-list-1280.png', 'Рисунок 4.6 — Робочий список звернень у кабінеті менеджера', width_cm=14.5)
    add_figure(doc, SCR_DIR / '08-manager-modal-detail.png', 'Рисунок 4.7 — Картка звернення з призначенням, зміною статусу та внутрішніми примітками', width_cm=14.5)
    add_figure(doc, SCR_DIR / '09-manager-sla-report.png', 'Рисунок 4.8 — Звіт про прострочені заявки та винятки виробничого календаря', width_cm=14.5)

    add_h2(doc, '4.3 Версіонування та публікація кодової бази')
    add_p(doc, 'Контроль версій ведеться у Git-репозиторії https://github.com/Zhyhan245/voltapro-service-inquiries під ліцензією MIT (основна гілка main, позначка релізу v1.0.0); конфігурація .env і каталог node_modules/ виключені правилами .gitignore.')

    add_h2(doc, '4.4 Кількісні метрики програмного коду')
    s = code_rep['summary']
    lf = s['largestFile']
    add_table(
        doc,
        'Таблиця 4.1 — Показники обсягу та якості кодової бази (tests/code-metrics.js)',
        ['Метрика', 'Зафіксоване значення', 'Коментар'],
        [
            ['Кількість JS-модулів', str(s['filesCount']), 'Чітке розділення серверної та клієнтської логіки'],
            ['Сумарна кількість рядків', f"{s['totalLines']} (змістовних: {s['nonEmptyLines']})", 'Власний код без автогенерованих фрагментів'],
            ['Кількість функцій', str(s['totalFunctions']), 'Компактні функції (у середньому ~19 рядків)'],
            ['Найбільший модуль', f"{lf['file']} ({lf['totalLines']} рядків)", 'У межах порогу 500 рядків; містить DTO-проєкції'],
            ['Тестове покриття (node:test)', '90,00 % рядків / 88,89 % функцій', 'Суттєво вище базової вимоги 70 %'],
            ['Аудит пакетів (npm audit)', '0 vulnerabilities', 'Усі 74 встановлені залежності безпечні']
        ],
        col_widths=[5.5, 4.5, 7.0]
    )

    # =========================================================================
    # РОЗДІЛ 5. ЗАБЕЗПЕЧЕННЯ ЯКОСТІ ВЕБСИСТЕМИ
    # =========================================================================
    add_h1(doc, '5 ТЕСТУВАННЯ ТА ВЕРИФІКАЦІЯ ЯКОСТІ')

    add_h2(doc, '5.1 Стратегія комплексної перевірки')
    add_p(doc, 'Верифікація реалізованого продукту охоплює п’ять послідовних рівнів: модульні тести (node:test), інтеграційні перевірки HTTP API разом із SQLite, автоматизовані браузерні сценарії (Playwright + Google Chrome), навантажувальні прогони (5, 10, 25 VU) та дослідження деградації швидкодії при зростанні бази до 10 000 записів.')

    add_h2(doc, '5.2 Реєстр контрольних сценаріїв')
    add_table(
        doc,
        'Таблиця 5.1 — Зведення тестових сценаріїв та результатів виконання',
        ['Код', 'Вимога', 'Вхідні дані та умови', 'Очікувана й отримана поведінка', 'Підсумок'],
        [
            ['TC-01', 'FR-01', 'Запит «монтаж» і фільтрація категорії «Аудит»', '200 OK, точний пошук кирилицею; коректна заглушка за відсутності збігів', 'Пройдено'],
            ['TC-02', 'FR-02', 'Надсилання форми з неповним контактом («А», «123»)', '400 VALIDATION_ERROR зі збереженням введених полів', 'Пройдено'],
            ['TC-03', 'FR-03', 'Дубльований запит із незмінним idempotencyKey', '200 OK (duplicate: true) без повторного запису в БД', 'Пройдено'],
            ['TC-04', 'FR-04, FR-05', 'Перевірка VP-2026-0001 / DEMO01 і хибного коду ZZZZ99', '200 OK без службових приміток; 404 для хибного коду', 'Пройдено'],
            ['TC-05', 'NFR-SEC-01', 'Звернення до /api/manager/inquiries без cookie сесії', '401 Unauthorized', 'Пройдено'],
            ['TC-06', 'FR-08', 'Доступ manager2 до заявки, закріпленої за manager1', '403 Forbidden без розкриття даних', 'Пройдено'],
            ['TC-07', 'FR-07', 'Спроба відхилити звернення з порожньою причиною', '400 REJECTION_REASON_REQUIRED', 'Пройдено'],
            ['TC-08', 'FR-08', 'Переведення COMPLETED → IN_PROGRESS різними ролями', '403 для MANAGER; 200 OK для SENIOR_MANAGER', 'Пройдено'],
            ['TC-09', 'FR-09', 'Обчислення SLA: 2026-12-24 + 2 роб. дн. (свято 25.12 + вихідні)', 'Підсумковий дедлайн 2026-12-29', 'Пройдено'],
            ['TC-10', 'FR-10', 'Формування звіту /api/manager/reports/overdue', 'Коректні лічильники активних, прострочених і «дедлайн сьогодні»', 'Пройдено']
        ],
        col_widths=[1.5, 2.0, 5.2, 6.3, 2.0]
    )

    add_h2(doc, '5.3 Результати модульних, інтеграційних і браузерних тестів')
    add_p(doc, 'Запуск node --test --experimental-test-coverage підтвердив успішне виконання всіх 7 тестових блоків із покриттям 90,00 % рядків та 88,89 % функцій. Автоматизований скрипт tests/e2e-screenshots.js у Google Chrome також виконав усі 10 наскрізних перевірок без жодного збою.')

    add_h2(doc, '5.4 Перевірка адаптивності верстки та доступності')
    add_p(doc, 'Тестування на контрольних ширинах 360, 768 і 1280 пікселів (рисунки 5.1, 5.2) засвідчило відсутність горизонтального зсуву (scrollWidth <= clientWidth), а час події DOMContentLoaded на локальному стенді вкладався в межі 30–60 мс.')
    add_figure(doc, SCR_DIR / '02-public-mobile-360.png', 'Рисунок 5.1 — Мобільна адаптивність публічного сайту на ширині 360 px', width_cm=8.0)
    add_figure(doc, SCR_DIR / '10-manager-tablet-768.png', 'Рисунок 5.2 — Планшетна адаптивність кабінету менеджера на ширині 768 px', width_cm=12.5)

    add_h2(doc, '5.5 Випробування під навантаженням і масштабуванням БД')
    add_p(doc, 'Показники пропускної здатності (tests/load-test.js) для 5, 10 та 25 віртуальних користувачів подано в таблиці 5.2 і на рисунку 5.3, а час відгуку SQLite на наборах 100, 1 000 та 10 000 записів (tests/volume-test.js) — у таблиці 5.3 та на рисунку 5.4.')
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

    add_h2(doc, '5.6 Виправлені дефекти та підсумок якості')
    add_table(
        doc,
        'Таблиця 5.4 — Перелік зафіксованих дефектів і результати регресійної перевірки',
        ['Ідентифікатор', 'Суть проблеми', 'Вимога', 'Вжите рішення та статус'],
        [
            ['DEF-01', 'Виклик path.resolve(":memory:") у src/config.js спотворював шлях у Windows під час запуску тестів у пам’яті.', 'NFR-RUN-01', 'Додано виняток для rawDbPath === ":memory:"; повторний запуск 7/7 тестів успішний (Закрито).'],
            ['DEF-02', 'Стандартна функція LOWER() у SQLite не переводила кириличні літери в нижній регістр при пошуку («монтаж»).', 'FR-01', 'Додано користувацьку SQL-функцію uk_lower у connection.js та автотест в integration.test.js (Закрито).']
        ],
        col_widths=[1.8, 6.2, 2.6, 6.4]
    )

    # =========================================================================
    # ВИСНОВКИ
    # =========================================================================
    add_h1(doc, 'ВИСНОВКИ')
    add_p(doc, 'У ході курсового проєктування розроблено та верифіковано корпоративний вебсайт підприємства «ВольтаПро Сервіс» із підсистемою обліку звернень (Варіант 5). Поставлені завдання виконано в повному обсязі: замовники отримали зрозумілий інтерфейс подання заявок і самостійного відстеження результату, а працівники — надійне робоче середовище з автоматичним контролем дедлайнів.')
    add_p(doc, 'Практичне значення роботи забезпечують чотирирівнева архітектура бекенду (Node.js, Express, SQLite у режимі WAL), адаптивна клієнтська частина (Bootstrap 5 і модулі ES6), захист від дубльованого надсилання форм, ізоляція службових нотаток від публічного перегляду, рольовий доступ (MANAGER та SENIOR_MANAGER) і календарне обчислення SLA у робочих днях зі звітом прострочень. Результати тестування підтвердили 90,00 % покриття коду, високу пропускну здатність (800,99 зап./с при 25 VU, p95 = 7,37 мс) та швидку роботу запитів на обсязі до 10 000 заявок.')
    add_p(doc, 'До меж поточної реалізації належать робота на локальному стенді без реального поштового шлюзу та використання файлової СКБД SQLite. Перспективними напрямами подальшого розвитку є підключення зовнішнього SMTP-сервісу повідомлень та міграція сховища на PostgreSQL.')

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
