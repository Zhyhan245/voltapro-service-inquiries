import json
from pathlib import Path
import docx
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

ROOT = Path(__file__).resolve().parent.parent
DIAG_DIR = ROOT / 'docs' / 'diagrams'
SCR_DIR = ROOT / 'evidence' / 'screenshots'
REP_DIR = ROOT / 'evidence' / 'reports'
OUT_DOCX = ROOT / 'docs' / 'Курсова_робота_Жиган_Д_О_АІ-245.docx'


def set_cell_shading(cell, fill_hex):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


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
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run()
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    fldChar1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
    instrText = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w'))
    fldChar2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
    fldChar3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)


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
        st.paragraph_format.space_before = Pt(12 if lvl > 1 else 0)
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


def add_h3(doc, text):
    return doc.add_heading(text, level=3)


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

    # Повтор заголовка на новій сторінці
    trPr = tbl.rows[0]._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

    for j, h in enumerate(headers):
        cell = tbl.rows[0].cells[j]
        set_cell_shading(cell, 'E5E7EB')
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
    fldChar1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
    instrText = parse_xml(r'<w:instrText %s xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText>' % nsdecls('w'))
    fldChar2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
    fldChar3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)


def build_document():
    doc = Document()
    configure_styles(doc)

    # Завантажуємо реальні метрики з evidence/reports/
    load_rep = json.loads((REP_DIR / 'load-report.json').read_text(encoding='utf-8'))['stages']
    vol_rep = json.loads((REP_DIR / 'volume-report.json').read_text(encoding='utf-8'))['results']
    code_rep = json.loads((REP_DIR / 'code-metrics.json').read_text(encoding='utf-8'))
    e2e_rep = json.loads((REP_DIR / 'e2e-ui-report.json').read_text(encoding='utf-8'))['checks']

    # =========================================================================
    # ТИТУЛЬНИЙ АРКУШ (Додаток В)
    # =========================================================================
    for line in [
        'Міністерство освіти і науки України',
        'Національний університет «Одеська політехніка»',
        'Навчально-науковий інститут комп’ютерних систем',
        'Кафедра інформаційних систем'
    ]:
        add_p(doc, line, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=2)

    add_p(doc, '', indent=0, space_after=48)
    add_p(doc, 'КУРСОВА РОБОТА', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=6)
    add_p(doc, 'з дисципліни «Веб-технології та веб-дизайн»', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=6)
    add_p(doc, 'на тему:', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=6)
    add_p(
        doc,
        '«Корпоративний сайт із керуванням зверненнями для підприємства «ВольтаПро Сервіс» (Варіант 5)»',
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        indent=0,
        space_after=28
    )

    add_p(doc, 'Спеціальність: 122 «Комп’ютерні науки»', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=4)
    add_p(doc, 'Освітня програма: «Комп’ютерні науки»', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=12)
    add_p(doc, 'Виконала: здобувачка групи АІ-245', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=2)
    add_p(doc, 'Жиган Дар’я Олександрівна', bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=4)
    add_p(doc, 'Підпис _____________________   Дата «____» ____________ 2026 р.', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=12)
    add_p(doc, 'Керівник: Червоненко Петро Петрович', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=4)
    add_p(doc, 'Підпис _____________________   Дата «____» ____________ 2026 р.', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=48)
    add_p(doc, 'Одеса — 2026', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=0)

    # =========================================================================
    # ЗАВДАННЯ НА КУРСОВУ РОБОТУ (Додаток Г)
    # =========================================================================
    doc.add_page_break()
    add_p(doc, 'Національний університет «Одеська політехніка»', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=2)
    add_p(doc, 'Кафедра інформаційних систем', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=6)
    add_p(doc, 'ЗАВДАННЯ НА КУРСОВУ РОБОТУ', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=2)
    add_p(doc, 'з дисципліни «Веб-технології та веб-дизайн»', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=8)

    add_p(doc, 'Здобувачка: Жиган Дар’я Олександрівна   Група: АІ-245', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=2)
    add_p(doc, 'Спеціальність / освітня програма: 122 «Комп’ютерні науки» / «Комп’ютерні науки»', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=2)
    add_p(doc, 'Керівник: Червоненко Петро Петрович', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=6)
    add_p(doc, '1. Тема роботи: Корпоративний сайт із керуванням зверненнями для підприємства «ВольтаПро Сервіс» (Варіант 5).', bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=4)

    add_table(
        doc,
        '2. Мета, користувачі та вихідні дані',
        ['Параметр', 'Зміст'],
        [
            ['Мета розроблення', 'Підвищити прозорість та оперативність опрацювання клієнтських звернень підприємства завдяки вебсистемі каталогу послуг, автоматичної реєстрації заявок, контролю строків SLA у робочих днях та рольового кабінету менеджера.'],
            ['Ролі користувачів', 'Гість / Клієнт; Менеджер; Старший менеджер.'],
            ['Основні дані та джерела', 'Послуга (services), Звернення (inquiries), Клієнтський контакт (client_contacts), Відповідальний працівник (employees), Історія статусів (inquiry_history), Виробничий календар (calendar_exceptions). Синтетичний набір seed-даних.'],
            ['Обов’язковий наскрізний сценарій', 'Вибір послуги -> подання звернення -> отримання номера й коду доступу -> вхід менеджера -> призначення відповідального -> зміна статусу та додавання коментарів -> перевірка статусу клієнтом без розкриття внутрішніх приміток.'],
            ['Погоджене ускладнення', 'Контроль строку опрацювання (SLA) у робочих днях за явно визначеним виробничим календарем та інтерактивний звіт про прострочені заявки.'],
            ['Поза межами роботи', 'Реальне надсилання SMS/Email-повідомлень, інтеграція з платіжними шлюзами та промислова експлуатація 24/7.']
        ],
        col_widths=[5.0, 12.0]
    )

    add_p(doc, '3. Строки: Дата видачі: 28.10.2026   Термін подання: 27.11.2026   Дата захисту: 27.11.2026   Формат: локальний запуск + демонстрація.', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=6)

    add_table(
        doc,
        '4. Обов’язкові функції та критерії приймання',
        ['Код / функція', 'Очікуваний результат і спосіб перевірки'],
        [
            ['FR-01 / Каталог послуг', 'Пошук і фільтрація активних послуг за категоріями; перевірка порожньої вибірки (API + UI тест).'],
            ['FR-02 / Подання звернення', 'Валідація контактів, збереження полів при 400, захист від дублювання через idempotencyKey (інтеграційний + E2E тест).'],
            ['FR-03 / Пошук заявки клієнтом', 'Перегляд статусу за номером і 6-символьним кодом доступу без внутрішніх приміток; 404 при помилці коду (API + E2E).'],
            ['FR-04 / Керування заявкою', 'Призначення відповідального, переходи «нова — в роботі — завершена / відхилена» з обов’язковою причиною відхилення (модульний + інтеграційний).'],
            ['FR-05 / Права доступу', 'Заборона доступу до чужої заявки (403) і повернення завершеної заявки без ролі старшого менеджера (403).'],
            ['FR-06 / Контроль SLA', 'Обчислення дедлайну за календарем робочих днів і формування звіту прострочених заявок (модульний + інтеграційний).'],
            ['NFR-01 / Продуктивність', 'p95 затримки API <= 800 мс при навантаженні до 25 VUs; перевірка навантаженням (load-test.js).'],
            ['NFR-02 / Збереження й запуск', 'Дані зберігаються у SQLite після перезапуску; проєкт розгортається за README у чистому каталозі.']
        ],
        col_widths=[5.0, 12.0]
    )

    add_p(doc, '5. Склад пояснювальної записки та графічні матеріали: 5 розділів за ДСТУ 3008:2015, UML-діаграма прецедентів, блок-схеми алгоритмів, карта станів UI, структурна схема, розгортання, шари, послідовність, ER-модель, діаграми класів/DTO, графіки тестування та скриншоти (у складі записки).', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=4)
    add_p(doc, 'Завдання видав: Червоненко П. П. (28.10.2026)     Завдання прийняла: Жиган Д. О. (28.10.2026)', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=0)

    # =========================================================================
    # АНОТАЦІЯ ТА SUMMARY
    # =========================================================================
    add_h1(doc, 'АНОТАЦІЯ')
    add_p(doc, 'Курсову роботу присвячено розробленню корпоративного вебсайту з системою керування зверненнями клієнтів для інженерно-сервісного підприємства «ВольтаПро Сервіс» (Варіант 5). Метою роботи є підвищення прозорості та оперативності опрацювання заявок клієнтів за рахунок автоматизованої реєстрації, регламентованого контролю строків виконання (SLA) у робочих днях та розмежування прав доступу персоналу.')
    add_p(doc, 'Вебсистему побудовано у вигляді модульного клієнт-серверного застосунку: інтерфейс користувача реалізовано засобами семантичного HTML5, CSS3, JavaScript (ES6+) та Bootstrap 5; серверну частину — на платформі Node.js з фреймворком Express; постійне сховище — на базі реляційної СКБД SQLite у режимі WAL із транзакціями та зовнішніми ключами. Реалізовано каталог послуг із пошуком і фільтрацією, форму звернення із захистом від повторного надсилання, пошук заявки за номером і кодом доступу без розкриття внутрішніх приміток, кабінет менеджера з рольовою моделлю (менеджер і старший менеджер), а також погоджене ускладнення — обчислення дедлайнів за виробничим календарем і звіт про прострочені заявки.')
    add_p(doc, 'Проведено модульне, інтеграційне (покриття коду 90,00%), наскрізне браузерне (Google Chrome, ширини 360, 768, 1280 пікселів), навантажувальне тестування (до 25 віртуальних користувачів, 800,99 запитів/с, p95 = 7,37 мс) та тестування обсягом даних до 10 000 записів.')
    add_p(doc, 'Ключові слова: КОРПОРАТИВНИЙ САЙТ, КЕРУВАННЯ ЗВЕРНЕННЯМИ, КОНТРОЛЬ SLA, РОБОЧІ ДНІ, NODE.JS, EXPRESS, SQLITE, BOOTSTRAP 5, ПРОБЛЕМНІ ДЕТАЛІ RFC 9457.', bold=True)

    add_h1(doc, 'SUMMARY')
    add_p(doc, 'This term paper is dedicated to the design and development of a corporate website with a client inquiry management system for the engineering and service enterprise "VoltaPro Service" (Variant 5). The objective of the project is to increase transparency and responsiveness in processing customer inquiries through automated registration, regulated SLA deadline tracking in business days, and role-based access control.')
    add_p(doc, 'The web system is implemented as a modular client-server application using semantic HTML5, CSS3, JavaScript (ES6+), and Bootstrap 5 for the user interface, Node.js with Express for the trusted server layer, and SQLite (WAL mode) with transactions and foreign keys for persistent storage. The implementation covers a searchable service catalog, an idempotent inquiry submission form, public status lookup by ticket number and access code (filtering out internal staff notes), a manager dashboard with role-based transitions, and an agreed extension: business-day SLA calculation using an explicit calendar and an overdue inquiries report.')
    add_p(doc, 'Unit, integration (90.00% line coverage), end-to-end browser testing (360, 768, 1280 px viewports), load testing (up to 25 virtual users, 800.99 req/s, p95 = 7.37 ms), and volume testing (up to 10,000 records) were executed and documented.')
    add_p(doc, 'Keywords: CORPORATE WEBSITE, INQUIRY MANAGEMENT, SLA TRACKING, BUSINESS DAYS, NODE.JS, EXPRESS, SQLITE, BOOTSTRAP 5, RFC 9457.', bold=True)

    # =========================================================================
    # ЗМІСТ ТА ПЕРЕЛІК СКОРОЧЕНЬ
    # =========================================================================
    add_h1(doc, 'ЗМІСТ')
    add_toc_field(doc)

    add_h1(doc, 'ПЕРЕЛІК УМОВНИХ ПОЗНАЧЕНЬ ТА СКОРОЧЕНЬ')
    for abbr, desc in [
        ('API', 'Application Programming Interface — програмний інтерфейс застосунку.'),
        ('CRUD', 'Create, Read, Update, Delete — базові операції роботи з даними.'),
        ('DTO', 'Data Transfer Object — об’єкт передавання даних між шарами або клієнтом і сервером.'),
        ('ER', 'Entity-Relationship — модель «сутність — зв’язок».'),
        ('HTTP / HTTPS', 'HyperText Transfer Protocol / Secure — протокол передавання гіпертексту.'),
        ('PK / FK', 'Primary Key / Foreign Key — первинний та зовнішній ключі реляційної таблиці.'),
        ('RBAC', 'Role-Based Access Control — керування доступом на основі ролей.'),
        ('REST', 'Representational State Transfer — архітектурний стиль мережевих інтерфейсів.'),
        ('SLA', 'Service Level Agreement — регламентований строк опрацювання звернення.'),
        ('SQL', 'Structured Query Language — мова структурованих запитів.'),
        ('UI', 'User Interface — користувацький інтерфейс.'),
        ('UML', 'Unified Modeling Language — уніфікована мова моделювання.'),
        ('WCAG', 'Web Content Accessibility Guidelines — настанови щодо доступності вебконтенту.')
    ]:
        add_bullet(doc, desc, bold_prefix=f'{abbr} — ')

    # =========================================================================
    # ВСТУП
    # =========================================================================
    add_h1(doc, 'ВСТУП')
    add_p(doc, 'Сучасні сервісні та інженерні підприємства щодня отримують десятки звернень від потенційних і постійних клієнтів через розрізнені канали: телефонні дзвінки, електронну пошту та месенджери. За відсутності єдиної вебсистеми обліку звернення губляться, клієнт не має прозорої інформації про стан розгляду свого запиту, а керівництво не може оперативно виявити прострочені задачі та перевантаження окремих менеджерів.')
    add_p(doc, 'Метою курсової роботи є підвищення прозорості та керованості процесу приймання й опрацювання клієнтських звернень підприємства «ВольтаПро Сервіс» шляхом створення корпоративного вебсайту з публічним каталогом послуг, онлайн-формою звернення, захищеним пошуком статусу заявки за її номером, рольовим кабінетом менеджера та автоматизованим контролем строків виконання (SLA) у робочих днях.')
    add_p(doc, 'Для досягнення поставленої мети в роботі розв’язано такі задачі: проаналізовано предметну область та щонайменше три релевантні аналоги; сформульовано перевірювані функціональні та нефункціональні вимоги; спроєктовано й обґрунтовано алгоритми розрахунку робочих днів SLA і контролю переходів статусів; розроблено архітектуру, реляційну модель даних у SQLite та контракти REST API; реалізовано серверну та адаптивну клієнтську частини вебсистеми; проведено комплексне тестування (модульне, інтеграційне, наскрізне, навантажувальне й обсягом даних) та підготовлено супровідну документацію.')
    add_p(doc, 'Об’єктом розроблення є процес приймання, розподілу та контролю виконання клієнтських звернень на корпоративному вебсайті. Предметом розроблення є моделі даних, бізнес-правила переходів статусів, алгоритми календарного контролю строків опрацювання та програмні засоби клієнт-серверної взаємодії.')

    # =========================================================================
    # РОЗДІЛ 1. АНАЛІЗ ПРЕДМЕТНОЇ ОБЛАСТІ ТА ВИМОГИ
    # =========================================================================
    add_h1(doc, '1 АНАЛІЗ ПРЕДМЕТНОЇ ОБЛАСТІ ТА ВИМОГИ')

    add_h2(doc, '1.1 Проблема, цільова аудиторія й аналіз аналогів')
    add_p(doc, 'До впровадження вебсистеми приймання звернень на умовному інженерно-сервісному підприємстві «ВольтаПро Сервіс» (проєктування, монтаж, аудит та обслуговування систем резервного живлення і сонячної енергетики) відбувалося через загальну електронну пошту та телефонні дзвінки. У такому процесі виникають типові проблеми: відсутність єдиного реєстраційного номера звернення; повторне надсилання однакових запитів клієнтами; неможливість самостійно перевірити поточний етап опрацювання; ризик випадкового пересилання клієнтові внутрішніх службових нотаток; відсутність об’єктивного обліку строків виконання з урахуванням вихідних і святкових днів.')
    add_p(doc, 'Цільова аудиторія системи поділяється на зовнішніх відвідувачів (приватних осіб і представників бізнесу, які обирають послугу та подають звернення) і внутрішній персонал підприємства (менеджерів відділів та старшого менеджера, які опрацьовують заявки та контролюють дотримання SLA).')
    add_p(doc, 'Для обґрунтування проєктних рішень проведено порівняльний аналіз трьох релевантних систем-аналогів за відкритою документацією (дата ознайомлення — 05.11.2026): Zendesk Support, Bitrix24 CRM та Jira Service Management. Результати порівняння зведено в таблиці 1.1.')

    add_table(
        doc,
        'Таблиця 1.1 — Порівняння аналогів предметної області',
        ['Критерій', 'Zendesk Support', 'Bitrix24 CRM', 'Jira Service Management', 'Вимога до власного рішення'],
        [
            ['Цільова аудиторія', 'Служби підтримки середнього й великого бізнесу [офіц. док.]', 'Відділи продажів і малий/середній бізнес', 'ІТ-команди та сервісні підрозділи (ITSM)', 'Мале/середнє підприємство: гість, менеджер, старший менеджер'],
            ['Основний процес', 'Тикет-система з чергами та макросами', 'Воронка лідів з вебформ у CRM', 'Портал запитів із процесами погодження', 'Публічний каталог -> заявка -> обробка -> контроль SLA'],
            ['Робота з помилками', 'Перевірено за док.: стандартна валідація форм', 'Не перевірено у платних тарифах', 'Перевірено за док.: обов’язкові поля workflow', 'Чіткі помилки RFC 9457, збереження полів форми, ідемпотентність'],
            ['Доступність і мобільний інтерфейс', 'Адаптивний вебпортал, підтримка WCAG', 'Адаптивний інтерфейс і мобільний застосунок', 'Адаптивний портал замовника', 'Адаптивність 360/768/1280 px, навігація з клавіатури'],
            ['Обмеження застосування', 'Платна підписка, надлишкова складність для малого сайту', 'Закрита екосистема, прив’язка до хмарних тарифів', 'Високий поріг входу для нетехнічних клієнтів', 'Легкий автономний стек, відкритий код, прозорий календар SLA']
        ],
        col_widths=[2.8, 3.5, 3.5, 3.5, 3.7]
    )

    add_p(doc, 'Висновок аналізу: комерційні платформи (Zendesk, Bitrix24, Jira Service Management) вимагають платної підписки та надмірні для корпоративного сайту окремого підприємства. Водночас доцільно запозичити з них найкращі практики: чітку модель переходів статусів, розмежування публічних відповідей і внутрішніх приміток персоналу, а також автоматичний розрахунок дедлайнів SLA за виробничим календарем.')

    add_h2(doc, '1.2 Типи користувачів і межі системи')
    add_p(doc, 'У системі визначено три ролі користувачів із чітко розмежованими повноваженнями (таблиця 1.2). Ролі не підміняються іменами окремих облікових записів, а перевіряються на сервері для кожного захищеного запиту та конкретного об’єкта.')

    add_table(
        doc,
        'Таблиця 1.2 — Розмежування повноважень ролей користувачів',
        ['Дія в системі', 'Гість / Клієнт', 'Менеджер (MANAGER)', 'Старший менеджер (SENIOR_MANAGER)'],
        [
            ['Перегляд каталогу послуг і пошук', 'Так', 'Так', 'Так'],
            ['Подання нового звернення', 'Так', 'Так', 'Так'],
            ['Пошук заявки за номером і кодом доступу (публічний вигляд)', 'Так', 'Так', 'Так'],
            ['Перегляд робочого списку звернень', 'Ні', 'Лише нерозподілені та свої', 'Усі звернення'],
            ['Призначення відповідального', 'Ні', 'Лише на себе', 'На будь-кого або зняття'],
            ['Зміна статусу «нова -> в роботі -> завершена / відхилена»', 'Ні', 'Лише для доступних заявок', 'Для будь-якої заявки'],
            ['Повернення завершеної/відхиленої заявки в роботу', 'Ні', 'Ні (403 Forbidden)', 'Так'],
            ['Перегляд і створення внутрішніх приміток', 'Ні', 'Для доступних заявок', 'Для будь-якої заявки'],
            ['Перегляд звіту про прострочені заявки (SLA)', 'Ні', 'За своїми/нерозподіленими', 'Повний звіт по підприємству']
        ],
        col_widths=[6.0, 3.0, 4.0, 4.0]
    )

    add_p(doc, 'Межі системи: до власної розробленої системи належать публічний вебсайт, кабінет менеджера, REST API, доменна логіка валідації/SLA та локальна база даних SQLite. Зовнішні залежності обмежені бібліотекою оформлення Bootstrap 5 (підключена локально для автономної роботи). Зовнішні поштові чи SMS-шлюзи свідомо винесено за межі курсової роботи — підтвердження реєстрації та історія спілкування відображаються безпосередньо у вебінтерфейсі.')

    add_h2(doc, '1.3 Функціональні та нефункціональні вимоги')
    add_p(doc, 'Функціональні вимоги згруповано за шістьма завершеними функціями (F-01 — F-06). Наскрізна вимога FR-02 «Реєстрація нового звернення» деталізована у форматі паспорта вимоги в таблиці 1.3.')

    add_table(
        doc,
        'Таблиця 1.3 — Деталізована функціональна вимога FR-02',
        ['Поле', 'Зміст'],
        [
            ['Ідентифікатор', 'FR-02'],
            ['Функція', 'F-02 — Приймання та підтвердження клієнтського звернення'],
            ['Актор', 'Гість / Клієнт'],
            ['Передумови', 'Обрана послуга існує в каталозі та є активною (is_active = 1).'],
            ['Вхідні дані', 'serviceId, contactName (2–100 симв.), companyName (опційно), phone (9–15 цифр), email, subject (4–160 симв.), message (10–2000 симв.), idempotencyKey.'],
            ['Основний результат', 'Створено рівно один контакт, одне звернення зі статусом NEW та початковий запис історії; повернено HTTP 201, номер VP-YYYY-NNNN, 6-символьний accessCode та dueDate.'],
            ['Повторний запит', 'За наявності того самого idempotencyKey дублікат не створюється; повертається HTTP 200 з даними вже створеного звернення (duplicate: true).'],
            ['Некоректні дані', 'При неповному контакті чи порушенні меж повертається HTTP 400 у форматі RFC 9457 (fieldErrors); стан сховища не змінюється, поля форми зберігаються.'],
            ['Критерії приймання', 'Успішне створення, повторне надсилання без дубліката, неповний контакт перевірені окремими тестами.']
        ],
        col_widths=[4.2, 12.8]
    )

    add_p(doc, 'Інші функціональні вимоги охоплюють: FR-01 — перегляд активного каталогу послуг із текстовим пошуком і фільтром за категоріями; FR-03 — публічний пошук заявки за номером та 6-символьним кодом доступу (внутрішні примітки з is_internal = 1 виключаються на рівні SQL-запиту); FR-04 — керування зверненням у кабінеті менеджера (призначення відповідального, переходи статусів «NEW -> IN_PROGRESS -> COMPLETED / REJECTED» з обов’язковою причиною відхилення від 5 символів); FR-05 — авторизація та заборона доступу до чужих заявок (HTTP 403) і повторного відкриття завершеної заявки звичайним менеджером (HTTP 403); FR-06 — обчислення дедлайну в робочих днях за календарем та звіт про прострочені заявки.')

    add_table(
        doc,
        'Таблиця 1.4 — Нефункціональні вимоги до вебсистеми',
        ['Код', 'Вимога', 'Спосіб перевірки та критерій'],
        [
            ['NFR-UI-01', 'Інтерфейс коректно відображається на ширинах 360, 768 і 1280 CSS-пікселів без горизонтального обрізання та доступний з клавіатури.', 'Автоматизована перевірка scrollWidth <= clientWidth у Chrome + ручний прохід клавіатурою (WCAG 2.2).'],
            ['NFR-DATA-01', 'Перезапуск сервера не призводить до втрати зареєстрованих звернень, історії статусів або сесій.', 'Створення запису, зупинка/запуск процесу Node.js, повторне читання з SQLite.'],
            ['NFR-PERF-01', 'На тестовому стенді p95 часу відповіді API не перевищує 800 мс при навантаженні до 25 віртуальних користувачів.', 'Відтворюваний навантажувальний прогін (tests/load-test.js).'],
            ['NFR-SEC-01', 'Паролі зберігаються у вигляді scrypt-хешу з сіллю; сесійні cookie мають прапорці HttpOnly та SameSite=Lax; внутрішні примітки недоступні клієнту.', 'Інспекція коду та негативні інтеграційні тести доступу.'],
            ['NFR-RUN-01', 'Система розгортається за інструкцією README у чистому каталозі без неописаних ручних змін.', 'Контрольний запуск команд npm install, migrate, seed, start.']
        ],
        col_widths=[2.6, 8.2, 6.2]
    )

    add_h2(doc, '1.4 Моделювання прецедентів та перевірка якості вимог')
    add_p(doc, 'Зв’язок акторів системи та основних прецедентів подано на UML-діаграмі прецедентів (рисунок 1.1). Межа прямокутника відокремлює можливості вебсистеми від зовнішніх акторів; назви прецедентів узгоджені з каталогом функцій F-01 — F-06.')
    add_figure(doc, DIAG_DIR / 'fig_1_1_use_cases.png', 'Рисунок 1.1 — UML-діаграма прецедентів користувачів системи «ВольтаПро Сервіс»', width_cm=15.0)
    add_p(doc, 'Усі сформульовані вимоги перевірено за шістьма критеріями якості (атомарність, перевірюваність, зрозумілість, повнота, несуперечливість, простежуваність). Кожна вимога має стабільний код і прив’язку до модулів реалізації та тест-кейсів у матриці трасування (Додаток А).')

    # =========================================================================
    # РОЗДІЛ 2. АЛГОРИТМІЧНЕ ЗАБЕЗПЕЧЕННЯ ВЕБСИСТЕМИ
    # =========================================================================
    add_h1(doc, '2 АЛГОРИТМІЧНЕ ЗАБЕЗПЕЧЕННЯ ВЕБСИСТЕМИ')

    add_h2(doc, '2.1 Алгоритм розрахунку дедлайну та оцінки стану SLA за виробничим календарем')
    add_p(doc, 'Ключовим алгоритмічним завданням погодженого ускладнення є коректне обчислення планової дати завершення опрацювання звернення (due_date) та кількості прострочених робочих днів без підміни робочих днів календарними. Алгоритм addBusinessDays реалізовано в модулі src/domain/calendarRules.js.')
    add_p(doc, 'Паспорт алгоритму: вхідні дані — дата реєстрації startDateStr (формат YYYY-MM-DD), норматив slaBusinessDays (ціле число від 1 до 30), асоціативний масив винятків календаря exceptionsMap (дата -> boolean). Вихід — дата дедлайну у форматі YYYY-MM-DD. Інваріант: повернена дата завжди є робочим днем за явно визначеним календарем, а кількість робочих днів на інтервалі (startDateStr; dueDate] дорівнює рівно slaBusinessDays. Захист від нескінченного циклу забезпечено обмежувачем guard = 3660 ітерацій (10 років).')

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

    add_p(doc, 'Оцінка складності алгоритму addBusinessDays: нехай d = businessDays (норматив SLA у робочих днях, d <= 30). Оскільки кожні 7 календарних днів містять 5 стандартних робочих днів, кількість ітерацій циклу не перевищує приблизно (7/5)*d + k, де k — кількість поспіль святкових винятків. Часова складність алгоритму становить O(d) (для обмеженого d <= 30 це практично O(1) за кількістю операцій), а просторова складність додаткової пам’яті — O(1), оскільки стан зводиться до лічильників cursor і remaining, а пошук у хеш-таблиці exceptionsMap виконується за O(1).')

    add_h2(doc, '2.2 Алгоритм перевірки переходів статусів та прав доступу')
    add_p(doc, 'Для запобігання некоректним змінам життєвого циклу звернення («нова — в роботі — завершена / відхилена») у модулі src/domain/inquiryRules.js реалізовано скінченний автомат допустимих переходів ALLOWED_TRANSITIONS. Додатково перевіряються рольові обмеження: повернення заявки зі статусів COMPLETED або REJECTED назад у IN_PROGRESS дозволено виключно ролі SENIOR_MANAGER, а перехід у REJECTED вимагає заповненої причини відхилення довжиною від 5 символів (після обрізання пробілів). Блок-схему алгоритму наведено на рисунку 2.2.')
    add_figure(doc, DIAG_DIR / 'fig_2_2_status_transition.png', 'Рисунок 2.2 — Блок-схема перевірки допустимості переходу статусу звернення', width_cm=14.5)
    add_p(doc, 'Теоретична складність перевірки переходу статусу становить O(1) за часом і O(1) за пам’яттю, оскільки кількість можливих станів (4 стани) є фіксованою константою. Проте довірена перевірка обов’язково виконується на сервері всередині транзакції: клієнтська форма лише дублює підказки інтерфейсу для зручності менеджера.')

    add_h2(doc, '2.3 Алгоритм ідемпотентної реєстрації та формування звіту прострочень')
    add_p(doc, 'При повторному натисканні кнопки надсилання форми або збої мережі клієнт передає той самий UUID-ключ idempotencyKey. Сервер спочатку перевіряє наявність запису за унікальним індексом idempotency_key. Якщо запис існує, нове звернення не створюється, а клієнту повертається наявна картка підтвердження. Якщо запис новий — у транзакції BEGIN IMMEDIATE атомарно створюються клієнтський контакт, запис звернення, формується людиночитаний номер VP-YYYY-NNNN та додається запис в історію inquiry_history.')
    add_p(doc, 'Формування звіту прострочень SLA (getOverdueReport) вибирає активні заявки (статуси NEW та IN_PROGRESS) із використанням складеного індексу idx_inquiries_status_due(status, due_date), обчислює стан кожної з них за evaluateSlaState та сортує прострочені заявки за спаданням кількості прострочених робочих днів. Для n активних заявок часова складність сортування становить O(n log n), а просторова — O(n). Експериментальна перевірка на наборах 100, 1 000 та 10 000 записів (розділ 5.5) підтвердила стабільний час виконання звіту в межах 0,53–53,69 мс.')

    # =========================================================================
    # РОЗДІЛ 3. ПРОЄКТУВАННЯ ВЕБСИСТЕМИ
    # =========================================================================
    add_h1(doc, '3 ПРОЄКТУВАННЯ ВЕБСИСТЕМИ')

    add_h2(doc, '3.1 Архетип та обмеження архітектури')
    add_p(doc, 'За класифікацією методичних вказівок застосунок належить до змішаного архетипу: клієнтський вебзастосунок на HTML5 + CSS3 + ES6+ JavaScript + Bootstrap 5 взаємодіє через REST API (JSON) із сервером застосунку на Node.js + Express, який керує реляційним сховищем SQLite. Архітектурний стиль — модульний моноліт: таке рішення обрано замість мікросервісів, оскільки спільний цикл розроблення, транзакційна цілісність у межах однієї БД та єдиний процес запуску забезпечують високу надійність і простоту тестування для обсягу курсової роботи.')

    add_table(
        doc,
        'Таблиця 3.1 — Паспорт архітектурних обмежень',
        ['Код і зміст', 'Джерело', 'Наслідок і перевірка'],
        [
            ['AC-01: HTML5, CSS3, JS, Bootstrap 5 (без важких SPA-фреймворків)', 'Погодження з керівником та вибір стека', 'Модульна структура ES6-скриптів у public/js/; локальна роздача Bootstrap 5 без залежності від зовнішнього CDN.'],
            ['AC-02: Node.js + Express + вбудований node:sqlite', 'Вибір технологічного стека (табл. 1.3)', 'Відсутність зовнішніх нативних залежностей компіляції; відтворюваний запуск через npm start на Windows 11.'],
            ['AC-03: Відсутність платних зовнішніх сервісів', 'Обмеження курсової роботи', 'Усі перевірки проводяться на локальному стенді; нотифікації фіксуються в журналі історії.'],
            ['AC-04: Підтримка Docker (Dockerfile, docker-compose.yml)', 'Вимоги до розгортання', 'Можливість запуску в контейнері зі збереженням файла БД у volume ./data.']
        ],
        col_widths=[5.5, 4.5, 7.0]
    )

    add_h2(doc, '3.2 Подання користувацького інтерфейсу (UI View)')
    add_p(doc, 'Інтерфейс системи складається з трьох сторінок (/, /status, /manager) та восьми змістовних станів, взаємозв’язок між якими показано на карті станів (рисунок 3.1). Для кожної форми реалізовано стани завантаження, порожньої вибірки, помилки валідації (зі збереженням введених користувачем полів), успіху та заборони доступу. Цільові ширини адаптивності — 360, 768 і 1280 CSS-пікселів; інформація про статус і прострочення SLA передається не лише кольором бейджа, а й текстом.')
    add_figure(doc, DIAG_DIR / 'fig_3_1_ui_state_map.png', 'Рисунок 3.1 — Карта екранів та станів користувацького інтерфейсу', width_cm=15.5)

    add_table(
        doc,
        'Таблиця 3.2 — Паспорти основних екранів інтерфейсу',
        ['Екран', 'Доступні ролі', 'Основні дані, стани та UI-патерни'],
        [
            ['UI-01: Каталог послуг (/)', 'Гість, Менеджер', 'GET /api/services; пошук і фільтр за категоріями; стани: завантаження, порожній результат, сітка карток.'],
            ['UI-02 / UI-03: Форма звернення й підтвердження (/)', 'Гість, Менеджер', 'POST /api/inquiries; інлайн-валідація, блокування кнопки при запиті, збереження полів при 400, картка підтвердження з номером і кодом.'],
            ['UI-04: Перевірка статусу (/status)', 'Гість / Клієнт', 'GET /api/inquiries/lookup; пошук за номером + кодом, відображення дедлайну SLA, публічної відповіді та історії без внутрішніх нотаток.'],
            ['UI-05: Вхід менеджера (/manager)', 'Неавтентифікований персонал', 'POST /api/auth/login; форма пошти/пароля, повідомлення 401 без розкриття деталей, перехід до панелі.'],
            ['UI-06 / UI-07 / UI-08: Кабінет, картка, звіт SLA (/manager)', 'Менеджер, Старший менеджер', 'Фільтрована таблиця з пагінацією, модальна картка з історією та внутрішніми примітками, вкладка звіту прострочених заявок.']
        ],
        col_widths=[4.5, 3.5, 9.0]
    )

    add_h2(doc, '3.3 Логічне подання та подання розгортання (Logical & Deployment Views)')
    add_p(doc, 'Логічне подання (рисунок 3.2) розмежовує недовірену клієнтську частину в браузері та довірену серверну частину Node.js + Express, яка єдиною має доступ до приватного сховища SQLite. Браузер не отримує прямого доступу до файлів БД або службових полів.')
    add_figure(doc, DIAG_DIR / 'fig_3_2_logical_view.png', 'Рисунок 3.2 — Структурна схема вебсистеми та її довіреної частини (Logical View)', width_cm=15.0)
    add_p(doc, 'Подання розгортання (рисунок 3.3) відображає фактичну конфігурацію навчального стенда: пристрій користувача з веббраузером Google Chrome взаємодіє по HTTP з портом 3000 сервера Node.js, який працює з локальним файлом data/voltapro.db (або в Docker-контейнері з примонтованим томом ./data).')
    add_figure(doc, DIAG_DIR / 'fig_3_3_deployment_view.png', 'Рисунок 3.3 — Схема розгортання навчальної вебсистеми (Deployment View)', width_cm=14.5)

    add_h2(doc, '3.4 Подання шарів, процесів і даних (Design, Process & Data Views)')
    add_p(doc, 'Серверну частину організовано у чотири логічні шари (рисунок 3.4): шар подання (маршрути та middleware), прикладний шар (сервіси координації сценаріїв), шар предметних правил (доменні інваріанти та календар SLA) та інфраструктурні адаптери (репозиторії SQLite та логер). Контролери маршрутів не містять SQL-запитів і бізнес-правил, а лише делегують виклики сервісам.')
    add_figure(doc, DIAG_DIR / 'fig_3_4_design_layers.png', 'Рисунок 3.4 — Розподіл відповідальності між логічними шарами (Design View)', width_cm=14.5)

    add_p(doc, 'На діаграмі послідовності (рисунок 3.5) показано архітектурно значущий процес створення звернення з урахуванням трьох гілок: ідемпотентного повтору, атомарного збереження у транзакції BEGIN IMMEDIATE та відмови при помилці валідації.')
    add_figure(doc, DIAG_DIR / 'fig_3_5_sequence_create.png', 'Рисунок 3.5 — Діаграма послідовності реєстрації звернення (Process View)', width_cm=15.5)

    add_p(doc, 'Модель даних у SQLite містить сім таблиць (рисунок 3.6): services, client_contacts, inquiries, inquiry_history, employees, sessions та calendar_exceptions. Увімкнено зовнішні ключі (PRAGMA foreign_keys = ON), режим журналювання WAL, обмеження CHECK на допустимі статуси й категорії, а також індекси для вибірок менеджера та звіту SLA. Доменні сутності та структури передавання даних (DTO) розмежовано (рисунок 3.7), щоб службові поля (наприклад, password_hash або access_code) не потрапляли у публічні відповіді API.')
    add_figure(doc, DIAG_DIR / 'fig_3_6_er_model.png', 'Рисунок 3.6 — ER-модель реляційної бази даних SQLite (Data View)', width_cm=15.5)
    add_figure(doc, DIAG_DIR / 'fig_3_7_domain_dto.png', 'Рисунок 3.7 — Діаграма основних доменних об’єктів та DTO', width_cm=15.5)

    add_table(
        doc,
        'Таблиця 3.3 — Словник даних основної сутності inquiries',
        ['Поле', 'Тип / обмеження', 'Призначення'],
        [
            ['id', 'INTEGER PRIMARY KEY AUTOINCREMENT', 'Внутрішній сурогатний ідентифікатор запису'],
            ['ticket_number', 'TEXT NOT NULL UNIQUE', 'Публічний номер заявки формату VP-YYYY-NNNN'],
            ['access_code', 'TEXT NOT NULL', '6-символьний код доступу для пошуку клієнтом'],
            ['service_id / contact_id', 'INTEGER NOT NULL FK', 'Зовнішні ключі на послугу та контакт клієнта'],
            ['assigned_employee_id', 'INTEGER FK (NULLABLE)', 'Відповідальний менеджер (NULL, якщо не призначено)'],
            ['status', 'TEXT CHECK IN (...)', 'Стан: NEW, IN_PROGRESS, COMPLETED, REJECTED'],
            ['rejection_reason', 'TEXT (NULLABLE)', 'Обов’язкова причина при статусі REJECTED'],
            ['idempotency_key', 'TEXT UNIQUE', 'Ключ запобігання дублюванню при повторному надсиланні'],
            ['sla_business_days / due_date', 'INTEGER / TEXT (YYYY-MM-DD)', 'Зафіксований норматив і дедлайн за календарем'],
            ['created_at / updated_at / closed_at', 'TEXT (ISO 8601)', 'Часові мітки створення, оновлення та закриття']
        ],
        col_widths=[4.8, 5.2, 7.0]
    )

    add_h2(doc, '3.5 Інтерфейси, безпека, операційне, інфраструктурне подання та стратегія доставки')
    add_p(doc, 'Взаємодія клієнта та сервера відбувається через REST API у форматі JSON (детальна специфікація OpenAPI 3.1 збережена у файлі docs/openapi.yaml). Усі помилки повертаються у передбачуваному форматі Problem Details (RFC 9457) із полями code, message, fieldErrors та requestId. Основні операції API наведено в таблиці 3.4.')

    add_table(
        doc,
        'Таблиця 3.4 — Каталог основних HTTP-операцій REST API',
        ['Метод і маршрут', 'Доступ', 'Призначення та типові відповіді'],
        [
            ['GET /api/health', 'Публічний', 'Перевірка стану процесу та SQLite (200 / 503)'],
            ['GET /api/services', 'Публічний', 'Каталог послуг із пошуком і фільтром категорії (200)'],
            ['POST /api/inquiries', 'Публічний', 'Реєстрація звернення з ідемпотентністю (201 / 200 / 400 / 404)'],
            ['GET /api/inquiries/lookup', 'Публічний (за кодом)', 'Пошук заявки за номером і кодом без внутрішніх нотаток (200 / 400 / 404)'],
            ['POST /api/auth/login | logout', 'Персонал', 'Вхід за scrypt-паролем, видача HttpOnly cookie (200 / 401 | 204)'],
            ['GET /api/manager/inquiries', 'Менеджер / Старший', 'Робочий список із фільтрами та пагінацією (200 / 401)'],
            ['PATCH /api/manager/inquiries/:id/assign', 'Менеджер / Старший', 'Призначення відповідального з перевіркою ролі (200 / 403)'],
            ['PATCH /api/manager/inquiries/:id/status', 'Менеджер / Старший', 'Зміна статусу за правилами переходів (200 / 400 / 403 / 409)'],
            ['POST /api/manager/inquiries/:id/comments', 'Менеджер / Старший', 'Додавання публічної відповіді або внутрішньої примітки (201 / 400 / 403)'],
            ['GET /api/manager/reports/overdue', 'Менеджер / Старший', 'Звіт про прострочені заявки та винятки календаря (200 / 401)']
        ],
        col_widths=[6.0, 3.8, 7.2]
    )

    add_p(doc, 'Безпека (Security View): паролі працівників хешуються алгоритмом scrypt (N=16384, r=8, p=1) з індивідуальною 16-байтною сіллю та перевіркою через crypto.timingSafeEqual; демонстраційний пароль має довжину 21 символ (відповідає орієнтиру NIST SP 800-63B-4 >= 15 символів). Сесії передаються в cookie з прапорцями HttpOnly, SameSite=Lax та обмеженим строком дії. Усі SQL-запити параметризовані (захист від SQL-ін’єкцій), на клієнті використовується функція escapeHtml (захист від XSS), а HTTP-заголовки захищено middleware Helmet.')
    add_p(doc, 'Операційне та інфраструктурне подання: сервер веде структурований JSON-журнал запитів із requestId без запису паролів і токенів, а маршрут /api/health перевіряє не лише роботу процесу Express, а й реальну доступність SQLite запитом SELECT 1. Усі експерименти виконано на фактичному стенді: ОС Microsoft Windows 11 Pro, процесор AMD Ryzen 7 9700X (8 ядер, 16 потоків), ОЗП 61,61 ГБ, Node.js v26.13.1, npm 11.20.0, браузер Google Chrome. Стратегія доставки передбачає розгортання з вихідного коду за README.md (або через Docker Compose) з автоматичною ініціалізацією схеми міграціями 001_init.sql та seed-даними.')

    # =========================================================================
    # РОЗДІЛ 4. РЕАЛІЗАЦІЯ ВЕБСИСТЕМИ
    # =========================================================================
    add_h1(doc, '4 РЕАЛІЗАЦІЯ ВЕБСИСТЕМИ')

    add_h2(doc, '4.1 Ключові модулі, контракти та шлях даних')
    add_p(doc, 'Програмну частину реалізовано у вигляді модульної структури (рисунок 4.1). Вхідний HTTP-запит проходить через Express, Helmet, cookieParser, attachSession, потрапляє у відповідний маршрутизатор (publicRoutes, authRoutes або managerRoutes), після чого викликається прикладний сервіс (inquiryService або authService). Сервіс звертається до чистих функцій доменних правил (calendarRules, inquiryRules, validationRules) та до репозиторіїв SQLite, що інкапсулюють SQL-запити й транзакції.')
    add_figure(doc, DIAG_DIR / 'fig_4_1_architecture_classes.png', 'Рисунок 4.1 — Ключові модулі та контракти серверної реалізації', width_cm=14.0)

    add_code_block(
        doc,
        'Лістинг 4.1 — Ключовий фрагмент транзакційного створення звернення (src/repositories/inquiryRepository.js)',
        """function createWithContactAndHistory({ contact, inquiry, historyComment }) {
  return runInTransaction(db, () => {
    const contactRes = db.prepare(
      `INSERT INTO client_contacts (contact_name, company_name, phone, email, created_at)
       VALUES (?, ?, ?, ?, ?)`
    ).run(contact.contactName, contact.companyName, contact.phone, contact.email, inquiry.createdAt);
    const contactId = Number(contactRes.lastInsertRowid);
    // Вставка звернення, генерація номера VP-YYYY-NNNN та запис в inquiry_history
    ...
    return findById(inquiryId);
  });
}"""
    )

    add_h2(doc, '4.2 Реалізація інтерфейсу користувача та демонстрація станів')
    add_p(doc, 'Клієнтську частину реалізовано на семантичному HTML5 із використанням локально підключеного Bootstrap 5 та ES6-модулів. На рисунку 4.2 наведено головну сторінку корпоративного сайту «ВольтаПро Сервіс» із каталогом послуг, фільтрацією та формою звернення. На рисунках 4.3 та 4.4 продемонстровано обробку некоректного введення (зі збереженням даних у формі) та екран підтвердження прийняття звернення з видачею номера, коду доступу й дедлайну SLA.')
    add_figure(doc, SCR_DIR / '01-public-desktop-1280.png', 'Рисунок 4.2 — Публічний сайт, каталог послуг та форма звернення (1280 px)', width_cm=14.5)
    add_figure(doc, SCR_DIR / '03-public-form-error-validation.png', 'Рисунок 4.3 — Помилка валідації при неповному контакті зі збереженням введених полів', width_cm=14.5)
    add_figure(doc, SCR_DIR / '04-public-inquiry-confirmed.png', 'Рисунок 4.4 — Картка підтвердження прийняття звернення з кодом доступу та дедлайном SLA', width_cm=14.5)

    add_p(doc, 'Перевірку статусу заявки клієнтом за номером VP-2026-0001 та кодом DEMO01 показано на рисунку 4.5: у публічній історії відображається офіційна відповідь менеджера, тоді як внутрішня службова примітка повністю відсутня. На рисунках 4.6, 4.7 та 4.8 показано робочий список менеджера з індикацією SLA, детальну картку звернення з візуальним розмежуванням внутрішніх приміток і вкладку звіту про прострочені заявки.')
    add_figure(doc, SCR_DIR / '05-status-lookup-public.png', 'Рисунок 4.5 — Сторінка публічної перевірки статусу заявки без витоку внутрішніх приміток', width_cm=14.5)
    add_figure(doc, SCR_DIR / '07-manager-list-1280.png', 'Рисунок 4.6 — Робочий список звернень у кабінеті менеджера', width_cm=14.5)
    add_figure(doc, SCR_DIR / '08-manager-modal-detail.png', 'Рисунок 4.7 — Детальна картка звернення з призначенням, зміною статусу та внутрішніми примітками', width_cm=14.5)
    add_figure(doc, SCR_DIR / '09-manager-sla-report.png', 'Рисунок 4.8 — Звіт про прострочені заявки та винятки виробничого календаря', width_cm=14.5)

    add_h2(doc, '4.3 Керування версіями, метрики коду та контроль якості реалізації')
    add_p(doc, 'Вихідний код проєкту розміщено у Git-репозиторії https://github.com/Zhyhan245/voltapro-service-inquiries під ліцензією MIT. Релізну версію для захисту зафіксовано тегом v1.0.0 (гілка main, без потрапляння node_modules/, .env та локальних файлів БД завдяки .gitignore).')
    add_p(doc, 'Для оцінки структури власного коду виконано автоматизований аналіз 26 JS-файлів у каталогах src/ та public/js/ (скрипт tests/code-metrics.js, без урахування сторонніх бібліотек із node_modules). Зведені показники наведено в таблиці 4.1.')

    s = code_rep['summary']
    lf = s['largestFile']
    add_table(
        doc,
        'Таблиця 4.1 — Метрики власного програмного коду',
        ['Показник', 'Виміряне значення', 'Інтерпретація'],
        [
            ['Кількість власних JS-файлів', str(s['filesCount']), 'Модульний поділ на шари (сервер + клієнт)'],
            ['Загальний обсяг коду (рядків)', f"{s['totalLines']} (непорожніх: {s['nonEmptyLines']})", 'Чистий код без шаблонних згенерованих файлів'],
            ['Кількість функцій', str(s['totalFunctions']), 'Середній розмір функції ~19 рядків'],
            ['Найбільший файл', f"{lf['file']} ({lf['totalLines']} рядків)", 'Містить DTO-мапери та сценарії; нижче порогу 500 рядків'],
            ['Покриття тестами (node:test)', '90,00% рядків / 88,89% функцій', 'Високе покриття доменних правил і сервісів'],
            ['Стан вразливостей залежностей', '0 vulnerabilities (npm audit)', 'Усі 74 пакети перевірено без відомих вразливостей']
        ],
        col_widths=[5.5, 4.5, 7.0]
    )

    add_table(
        doc,
        'Таблиця 4.2 — Контроль якості реалізації',
        ['Питання контролю', 'Статус', 'Доказ'],
        [
            ['Чи завершений основний сценарій?', 'Так', 'Наскрізний інтеграційний та E2E тести зі збереженням у SQLite'],
            ['Чи перевіряє сервер права на об’єкт?', 'Так', 'Негативні тести 401 (без сесії) та 403 (чужа заявка / повторне відкриття)'],
            ['Чи зберігається цілісність даних?', 'Так', 'Зовнішні ключі, CHECK-обмеження, транзакції BEGIN IMMEDIATE, idempotency_key'],
            ['Чи винесено секрети з коду?', 'Так', 'Конфігурація через змінні середовища та .env.example без чинних паролів'],
            ['Чи можна запустити проєкт у чистому середовищі?', 'Так', 'Перевірено послідовністю npm install -> migrate -> seed -> start']
        ],
        col_widths=[6.0, 2.0, 9.0]
    )

    # =========================================================================
    # РОЗДІЛ 5. ЗАБЕЗПЕЧЕННЯ ЯКОСТІ ВЕБСИСТЕМИ
    # =========================================================================
    add_h1(doc, '5 ЗАБЕЗПЕЧЕННЯ ЯКОСТІ ВЕБСИСТЕМИ')

    add_h2(doc, '5.1 План тестування та протокол тест-кейсів')
    add_p(doc, 'Програма перевірки якості охоплює п’ять видів випробувань: модульне (node:test), інтеграційне (HTTP API + SQLite), наскрізне браузерне (Playwright + Google Chrome), навантажувальне (5, 10, 25 віртуальних користувачів) та тестування обсягом даних (100, 1 000, 10 000 записів). Усі тести виконано на фінальній версії коду.')

    add_table(
        doc,
        'Таблиця 5.1 — Протокол виконання ключових тест-кейсів',
        ['Код тесту', 'Сценарій і вимога', 'Очікуваний результат', 'Статус'],
        [
            ['TC-01', 'Вибір послуги і фільтрація (FR-01)', 'Відображаються послуги обраної категорії; при відсутності збігів — порожній стан', 'Пройдено'],
            ['TC-02', 'Неповний контакт у формі (FR-02)', 'HTTP 400, інлайн-помилки біля полів, введені дані не очищаються', 'Пройдено'],
            ['TC-03', 'Повторне надсилання з idempotencyKey (FR-02)', 'HTTP 200 (duplicate: true), другий запис у БД не створюється', 'Пройдено'],
            ['TC-04', 'Публічний пошук за номером і кодом (FR-03)', 'HTTP 200, внутрішня примітка відсутня у відповіді; при невірному коді — 404', 'Пройдено'],
            ['TC-05', 'Неавторизований запит до кабінету (FR-05)', 'HTTP 401 Unauthorized на /api/manager/inquiries', 'Пройдено'],
            ['TC-06', 'Доступ manager2 до заявки manager1 (FR-05)', 'HTTP 403 Forbidden, дані чужої заявки не розкриті', 'Пройдено'],
            ['TC-07', 'Відхилення заявки без причини (FR-04)', 'HTTP 400 REJECTION_REASON_REQUIRED, статус не змінюється', 'Пройдено'],
            ['TC-08', 'Повернення завершеної заявки в роботу (FR-05)', 'Для MANAGER — HTTP 403; для SENIOR_MANAGER — HTTP 200', 'Пройдено'],
            ['TC-09', 'Розрахунок SLA через вихідні й свята (FR-06)', '2026-12-24 + 2 роб. дн. (з пропуском 25.12 та сб/нд) = 2026-12-29', 'Пройдено'],
            ['TC-10', 'Звіт про прострочені заявки (FR-06)', 'Повертає точні лічильники та перелік прострочених заявок', 'Пройдено']
        ],
        col_widths=[1.8, 5.2, 7.5, 2.5]
    )

    add_h2(doc, '5.2 Модульні, інтеграційні та наскрізні перевірки')
    add_p(doc, 'Запуск команди node --test --experimental-test-coverage tests/*.test.js підтвердив проходження усіх 7 тестових блоків (6 модульних та 1 комплексного інтеграційного) за 259,5 мс із 0 помилок та сумарним покриттям рядків коду 90,00% (покриття гілок — 72,87%, функцій — 88,89%). Усі 10 наскрізних браузерних перевірок у Google Chrome (tests/e2e-screenshots.js) також завершилися успішно.')

    add_h2(doc, '5.3 Адаптивність, доступність та швидкодія інтерфейсу')
    add_p(doc, 'Перевірку адаптивності виконано у Google Chrome на контрольних ширинах 360, 768 та 1280 CSS-пікселів (рисунки 5.1 та 5.2). На жодній ширині не виявлено горизонтального переповнення (scrollWidth <= clientWidth). Для доступності за WCAG 2.2 перевірено навігацію з клавіатури (Tab / Enter / Escape), наявність skip-link, видимий контур фокусу (:focus-visible), підписи label для всіх полів та дублювання кольорових бейджів текстом. Час DOMContentLoaded головної сторінки у локальному середовищі склав близько 30–60 мс.')
    add_figure(doc, SCR_DIR / '02-public-mobile-360.png', 'Рисунок 5.1 — Перевірка мобільної адаптивності публічного сайту на ширині 360 px', width_cm=8.0)
    add_figure(doc, SCR_DIR / '10-manager-tablet-768.png', 'Рисунок 5.2 — Перевірка планшетної адаптивності кабінету менеджера на ширині 768 px', width_cm=12.5)

    add_h2(doc, '5.4 Навантажувальне тестування')
    add_p(doc, 'Навантажувальне тестування проведено скриптом tests/load-test.js на локальному стенді (Windows 11 Pro, AMD Ryzen 7 9700X, 61,61 ГБ ОЗП) після 1 с прогріву трьома послідовними етапами по 5 с для 5, 10 та 25 віртуальних користувачів (змішаний профіль: каталог послуг, публічний пошук заявки та список менеджера). Цільовий поріг: p95 <= 800 мс та частка помилок < 1%. Фактичні результати зведено в таблиці 5.2 та на рисунку 5.3.')

    add_table(
        doc,
        'Таблиця 5.2 — Результати навантажувального тестування (tests/load-test.js)',
        ['Віртуальні користувачі', 'Запитів за 5 с', 'RPS (зап./с)', 'p50, мс', 'p95, мс', 'p99, мс', 'Помилки, %', 'RSS RAM, МБ'],
        [
            [str(st['virtualUsers']), str(st['totalRequests']), str(st['rps']), str(st['p50Ms']), str(st['p95Ms']), str(st['p99Ms']), str(st['errorRatePercent']), str(st['rssMemoryMb'])]
            for st in load_rep
        ],
        col_widths=[2.4, 2.2, 2.2, 1.8, 1.8, 1.8, 2.0, 2.8]
    )
    add_figure(doc, DIAG_DIR / 'fig_5_1_load_chart.png', 'Рисунок 5.3 — Залежність затримки відповіді та пропускної здатності від навантаження', width_cm=15.5)
    add_p(doc, 'Висновок: на тестовому стенді система стабільно обробила 4 025 запитів при 25 віртуальних користувачах із пропускною здатністю 800,99 запитів/с, нульовою часткою помилок та затримкою p95 = 7,37 мс, що суттєво нижче визначеного порогу 800 мс.')

    add_h2(doc, '5.5 Тестування обсягом даних та журнал дефектів')
    add_p(doc, 'Тестування обсягом даних виконано скриптом tests/volume-test.js шляхом генерації правдоподібних синтетичних звернень у трьох вибірках: 100, 1 000 та 10 000 записів (таблиця 5.3, рисунок 5.4). За допомогою EXPLAIN QUERY PLAN підтверджено використання складеного індексу idx_inquiries_status_due для вибірки прострочених заявок.')

    add_table(
        doc,
        'Таблиця 5.3 — Результати тестування обсягом даних (tests/volume-test.js)',
        ['Обсяг БД (записів)', 'Фільтрований список (avg / p95, мс)', 'Текстовий пошук (avg / p95, мс)', 'Звіт прострочень SLA (avg / p95, мс)'],
        [
            [
                str(r['recordsCount']),
                f"{r['listFiltered']['avgMs']} / {r['listFiltered']['p95Ms']}",
                f"{r['searchByText']['avgMs']} / {r['searchByText']['p95Ms']}",
                f"{r['overdueReport']['avgMs']} / {r['overdueReport']['p95Ms']}"
            ]
            for r in vol_rep
        ],
        col_widths=[3.5, 4.5, 4.5, 4.5]
    )
    add_figure(doc, DIAG_DIR / 'fig_5_2_volume_chart.png', 'Рисунок 5.4 — Час виконання запитів до SQLite залежно від обсягу даних', width_cm=14.5)

    add_p(doc, 'Журнал дефектів: під час тестування та ручної перевірки виявлено й усунено два дефекти. DEF-01: функція path.resolve(":memory:") у модулі src/config.js перетворювала рядок ":memory:" для тестової БД у пам’яті на некоректний шлях Windows, що спричиняло ERR_SQLITE_ERROR; виправлено явною перевіркою rawDbPath === ":memory:". DEF-02: вбудована функція LOWER() у SQLite не перетворює не-ASCII (кириличні) символи в нижній регістр, через що пошук за словом «монтаж» не знаходив записи; виправлено реєстрацією власної SQL-функції uk_lower у src/db/connection.js і додаванням перевірки кириличного пошуку в tests/integration.test.js.')

    # =========================================================================
    # ВИСНОВКИ
    # =========================================================================
    add_h1(doc, 'ВИСНОВКИ')
    add_p(doc, 'У ході виконання курсової роботи спроєктовано, реалізовано та всебічно перевірено корпоративний вебсайт із системою керування зверненнями для інженерно-сервісного підприємства «ВольтаПро Сервіс» (Варіант 5). Поставленої мети досягнуто: створено працездатне рішення, яке забезпечує прозорий процес приймання заявок, їхнього життєвого циклу та контролю строків виконання.')
    add_p(doc, 'Основні практичні результати роботи: розроблено серверну частину на Node.js + Express із чітким поділом на шари подання, прикладних сервісів, доменних правил та репозиторіїв SQLite; реалізовано адаптивний інтерфейс на HTML5, CSS3, JavaScript і Bootstrap 5; забезпечено ідемпотентне створення звернень, захищений пошук заявки клієнтом без витоку внутрішніх приміток персоналу, а також рольову модель доступу (Гість, Менеджер, Старший менеджер). Повністю виконано погоджене ускладнення — обчислення дедлайнів SLA у робочих днях за виробничим календарем та інтерактивний звіт про прострочені заявки.')
    add_p(doc, 'Якість реалізації підтверджено автоматизованими перевірками: модульні та інтеграційні тести показали 90,00% покриття рядків коду без жодної помилки; наскрізні перевірки у Google Chrome підтвердили адаптивність на ширинах 360, 768 та 1280 пікселів; навантажувальний тест зафіксував 800,99 запитів/с при 25 віртуальних користувачах (p95 = 7,37 мс), а тест обсягом на 10 000 записів підтвердив швидкодію пошуку та звітності в межах 10,6–53,7 мс.')
    add_p(doc, 'До обмежень поточної навчальної версії належать використання локального стенда без реального SMTP/SMS-шлюзу та однопроцесний запуск SQLite. Напрямами подальшого розвитку системи можуть бути інтеграція з чергою сповіщень електронною поштою, перенесення сховища на PostgreSQL при зростанні навантаження та впровадження двофакторної автентифікації для старших менеджерів.')

    # =========================================================================
    # ПЕРЕЛІК ДЖЕРЕЛ
    # =========================================================================
    add_h1(doc, 'ПЕРЕЛІК ДЖЕРЕЛ')
    sources = [
        '1. Арсірій О. О., Червоненко П. П. Методичні вказівки до курсової роботи з дисципліни «Веб-технології та веб-дизайн». Одеса: ОНПУ, 2017. 15 с.',
        '2. Методичні вказівки до виконання курсової роботи з дисципліни «Веб-технології та веб-дизайн». Об’єднана кафедральна редакція. Одеса: НУ «Одеська політехніка», 2026. 92 с.',
        '3. ДСТУ 3008:2015. Інформація та документація. Звіти у сфері науки і техніки. Структура та правила оформлювання. Київ: ДП «УкрНДНЦ», 2016.',
        '4. ДСТУ 8302:2015. Інформація та документація. Бібліографічне посилання. Загальні положення та правила складання. Київ: ДП «УкрНДНЦ», 2016.',
        '5. WHATWG. HTML Living Standard. URL: https://html.spec.whatwg.org/multipage/ (дата звернення: 05.11.2026).',
        '6. MDN Web Docs. CSS Grid Layout & Accessibility. URL: https://developer.mozilla.org/en-US/docs/Web/CSS (дата звернення: 05.11.2026).',
        '7. Bootstrap Documentation. Version 5.3. URL: https://getbootstrap.com/docs/5.3/ (дата звернення: 05.11.2026).',
        '8. Node.js Documentation. SQLite (node:sqlite) & Test Runner. URL: https://nodejs.org/docs/latest/api/ (дата звернення: 10.11.2026).',
        '9. Express.js Documentation. Routing and Middleware. URL: https://expressjs.com/ (дата звернення: 10.11.2026).',
        '10. IETF. RFC 9110: HTTP Semantics. 2022. URL: https://www.rfc-editor.org/rfc/rfc9110.html (дата звернення: 10.11.2026).',
        '11. IETF. RFC 9457: Problem Details for HTTP APIs. 2023. URL: https://www.rfc-editor.org/rfc/rfc9457.html (дата звернення: 10.11.2026).',
        '12. OWASP Foundation. OWASP Top 10:2025 & Password Storage Cheat Sheet. URL: https://owasp.org/ (дата звернення: 12.11.2026).',
        '13. NIST. SP 800-63B-4: Digital Identity Guidelines. 2025. URL: https://pages.nist.gov/800-63-4/sp800-63b.html (дата звернення: 12.11.2026).',
        '14. W3C. Web Content Accessibility Guidelines (WCAG) 2.2. URL: https://www.w3.org/TR/WCAG22/ (дата звернення: 15.11.2026).',
        '15. Object Management Group. Unified Modeling Language, version 2.5.1. URL: https://www.omg.org/spec/UML/2.5.1 (дата звернення: 15.11.2026).'
    ]
    for src in sources:
        add_p(doc, src, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=4)

    # =========================================================================
    # ДОДАТКИ
    # =========================================================================
    add_h1(doc, 'ДОДАТОК А. Матриця відповідності вимог і результатів')
    add_table(
        doc,
        'Таблиця А.1 — Матриця трасування вимог, реалізації та тестів',
        ['Вимога', 'Модуль реалізації', 'Тест / перевірка', 'Доказ', 'Статус'],
        [
            ['FR-01 Каталог послуг', 'serviceRepository.js, public.js', 'TC-01, e2e-screenshots.js', 'Рис. 4.2, e2e-ui-report.json', 'Виконано'],
            ['FR-02 Подання та ідемпотентність', 'inquiryService.js, validationRules.js', 'TC-02, TC-03, integration.test.js', 'Рис. 4.3, 4.4, тест-протокол', 'Виконано'],
            ['FR-03 Публічний пошук без приміток', 'inquiryService.js (lookupPublicInquiry)', 'TC-04, integration.test.js', 'Рис. 4.5, тест-протокол', 'Виконано'],
            ['FR-04 Статуси та причина відмови', 'inquiryRules.js, manager.js', 'TC-07, unit.test.js', 'Рис. 4.7, unit-протокол', 'Виконано'],
            ['FR-05 Права доступу ролей', 'auth.js, inquiryRules.js', 'TC-05, TC-06, TC-08', 'Негативні тести 401/403', 'Виконано'],
            ['FR-06 Календар SLA та звіт', 'calendarRules.js, inquiryService.js', 'TC-09, TC-10, unit/integration', 'Рис. 4.8, volume-report.json', 'Виконано'],
            ['NFR-UI-01 Адаптивність', 'styles.css, Bootstrap 5', 'e2e-screenshots.js (360/768/1280)', 'Рис. 5.1, 5.2', 'Виконано'],
            ['NFR-PERF-01 Продуктивність', 'SQLite WAL, індекси', 'load-test.js, volume-test.js', 'Рис. 5.3, 5.4, load-report.json', 'Виконано']
        ],
        col_widths=[3.2, 4.2, 3.8, 3.6, 2.2]
    )

    add_h1(doc, 'ДОДАТОК Б. Зведений протокол приймання та контрольний список')
    add_table(
        doc,
        'Таблиця Б.1 — Зведений протокол приймання вебсистеми',
        ['Група перевірок', 'Пройдено', 'Не пройдено', 'Заблоковано', 'Коментар'],
        [
            ['Бізнес-сценарії та правила', '7', '0', '0', 'Модульні та інтеграційні тести (90,00% coverage)'],
            ['Права та захист даних', '5', '0', '0', 'Перевірено 401/403, scrypt та фільтрацію приміток'],
            ['Інтерфейс і доступність', '10', '0', '0', 'Chrome E2E на 360, 768, 1280 px без обрізання'],
            ['Швидкодія й обсяг даних', '6', '0', '0', 'Етапи 5/10/25 VUs та обсяги 100/1000/10000'],
            ['Запуск і документація', '3', '0', '0', 'migrate -> seed -> start за README.md']
        ],
        col_widths=[4.5, 2.0, 2.4, 2.4, 5.7]
    )

    add_table(
        doc,
        'Таблиця Б.2 — Облік використання допоміжних засобів (ШІ)',
        ['Засіб і дата', 'Задача', 'Спосіб перевірки результату'],
        [
            ['AI-асистент (листопад 2026)', 'Генерація початкової структури модулів, тестових seed-даних і діаграм', 'Перевірено ручним переглядом коду, модульними/інтеграційними тестами, запуском у Chrome та нагрузочним тестуванням.']
        ],
        col_widths=[4.2, 5.8, 7.0]
    )

    add_h1(doc, 'ДОДАТОК В. Супровідна документація (інструкції користувача, запуску, адміністратора й API)')
    add_h2(doc, 'В.1 Інструкція користувача')
    add_p(doc, 'Для подання звернення відкрийте головну сторінку (/), скористайтеся пошуком або фільтром за категорією в каталозі послуг, натисніть «Обрати» на потрібній картці, заповніть контактні дані та опис задачі і натисніть «Надіслати звернення». Обов’язково збережіть виданий номер заявки (VP-YYYY-NNNN) та 6-символьний код доступу. Для перевірки стану опрацювання перейдіть на сторінку «Перевірити заявку» (/status) і введіть збережені реквізити.')
    add_h2(doc, 'В.2 Інструкція зі збирання та запуску')
    add_p(doc, 'Вимоги: Node.js v22.5+ (перевірено на v26.13.1), npm 11+. У каталозі проєкту виконайте послідовно: npm install; cp .env.example .env; npm run migrate; npm run seed; npm start. Після запуску сервер доступний на http://localhost:3000. Для зупинки натисніть Ctrl+C. Альтернативний запуск у контейнері: docker compose up --build -d.')
    add_h2(doc, 'В.3 Інструкція адміністратора')
    add_p(doc, 'Конфігураційні параметри задаються у файлі .env (PORT, DB_PATH, SESSION_TTL_HOURS). База даних зберігається у файлі data/voltapro.db. Для резервного копіювання достатньо зупинити процес сервера (або виконати sqlite3 .backup) і скопіювати файл БД. Демонстраційні облікові записи створюються скриптом npm run seed (senior@voltapro.ua, manager1@voltapro.ua, manager2@voltapro.ua). Попередження: видалення файла data/voltapro.db призведе до повного очищення накопичених звернень.')
    add_h2(doc, 'В.4 Опис програмного інтерфейсу (API)')
    add_p(doc, 'Повний опис точок доступу, схем вхідних/вихідних даних та кодів відмов наведено у специфікації OpenAPI 3.1 (файл docs/openapi.yaml) та узагальнено в таблиці 3.4 розділу 3.5.')

    doc.save(OUT_DOCX)
    print('Збережено DOCX:', OUT_DOCX)


if __name__ == '__main__':
    build_document()
