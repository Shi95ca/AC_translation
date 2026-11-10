#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Утилита подготовки обезличенных файлов для публикации проекта на GitHub:
1. Очищает шаблон Word (template.docx) от персональных данных, лишних таблиц и скрытых метаданных (core.xml / app.xml).
2. Генерирует демонстрационную карточку студента sample_card.xlsx (Иванов Иван Иванович) со структурой АИС «Студент».
"""

import os
import re
import sys
import copy
import zipfile
import tempfile
import shutil
from datetime import datetime, timezone

import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Настройки обезличенного подписанта (при желании можно заменить на нейтральное имя)
SIGNER_NAME = 'Evgeny Kurov'
Placeholder_INSPECTOR = 'I.I. Petrova'
PLACEHOLDER_DATE = '01.01.2026'


def fix_doc_defaults_and_styles(doc):
    """Приводит базовые интервалы шаблона к компактному виду (0 пт после абзаца, одинарный интервал, Times New Roman)."""
    styles_elm = doc.styles.element
    doc_defs = styles_elm.find(qn('w:docDefaults'))
    if doc_defs is not None:
        ppr_def = doc_defs.find(qn('w:pPrDefault'))
        if ppr_def is not None:
            ppr = ppr_def.find(qn('w:pPr'))
            if ppr is not None:
                sp = ppr.find(qn('w:spacing'))
                if sp is None:
                    sp = OxmlElement('w:spacing')
                    ppr.append(sp)
                sp.set(qn('w:before'), '0')
                sp.set(qn('w:after'), '0')
                sp.set(qn('w:line'), '240')
                sp.set(qn('w:lineRule'), 'auto')
        rpr_def = doc_defs.find(qn('w:rPrDefault'))
        if rpr_def is not None:
            rpr = rpr_def.find(qn('w:rPr'))
            if rpr is not None:
                rfonts = rpr.find(qn('w:rFonts'))
                if rfonts is not None:
                    for attr in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
                        rfonts.set(qn(attr), 'Times New Roman')
                    for t_attr in ('w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme'):
                        if qn(t_attr) in rfonts.attrib:
                            del rfonts.attrib[qn(t_attr)]

    for st_name in ('Normal', 'normal'):
        if st_name in doc.styles:
            st = doc.styles[st_name]
            st.paragraph_format.space_before = Pt(0)
            st.paragraph_format.space_after = Pt(0)
            st.paragraph_format.line_spacing = 1.0
            st.font.name = 'Times New Roman'
            st.font.size = Pt(11)


def clean_paragraph(p, line_spacing=1.0, keep_with_next=False, align=None):
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = line_spacing
    pf.keep_with_next = keep_with_next
    if align is not None:
        p.alignment = align
    if p._p.pPr is not None:
        rpr = p._p.pPr.find(qn('w:rPr'))
        if rpr is not None:
            for tag in ('w:highlight', 'w:sz', 'w:szCs'):
                el = rpr.find(qn(tag))
                if el is not None:
                    rpr.remove(el)


def write_run(p, text, bold=False, italic=False, size_pt=11):
    run = p.add_run(str(text))
    run.bold = bold
    run.italic = italic
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size_pt)
    run.font.color.rgb = RGBColor(0, 0, 0)
    rpr = run._r.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.insert(0, rfonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rfonts.set(qn(attr), 'Times New Roman')
    return run


def set_cell_value(cell, text, bold=False, italic=False):
    for extra_p in cell.paragraphs[1:]:
        cell._tc.remove(extra_p._p)
    p = cell.paragraphs[0]
    for r in list(p.runs):
        p._p.remove(r._r)
    clean_paragraph(p, line_spacing=1.0)
    if text != '':
        write_run(p, text, bold=bold, italic=italic)


def scrub_docx_zip_metadata(docx_path):
    """Удаляет скрытые следы редактирования из docProps/app.xml и закладки из колонтитулов."""
    tmp_fd, tmp_path = tempfile.mkstemp(suffix='.docx')
    os.close(tmp_fd)
    try:
        with zipfile.ZipFile(docx_path, 'r') as zin, zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                raw = zin.read(item.filename)
                if item.filename == 'docProps/app.xml':
                    txt = raw.decode('utf-8')
                    txt = re.sub(r'<TotalTime>\d+</TotalTime>', '<TotalTime>0</TotalTime>', txt)
                    txt = re.sub(r'<Company>.*?</Company>', '<Company></Company>', txt)
                    raw = txt.encode('utf-8')
                elif item.filename.startswith('word/header'):
                    txt = raw.decode('utf-8')
                    txt = re.sub(r'<w:bookmarkStart[^/]*/>', '', txt)
                    txt = re.sub(r'<w:bookmarkEnd[^/]*/>', '', txt)
                    raw = txt.encode('utf-8')
                zout.writestr(item, raw)
        shutil.move(tmp_path, docx_path)
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


def sanitize_template_docx(source_docx, target_docx):
    """Очищает шаблон Word от всех личных данных, оставляя эталонную структуру таблиц и колонтитулов."""
    doc = docx.Document(source_docx)
    fix_doc_defaults_and_styles(doc)

    # 1. Заголовок документа
    if len(doc.paragraphs) > 0:
        p0 = doc.paragraphs[0]
        for r in list(p0.runs):
            p0._p.remove(r._r)
        clean_paragraph(p0)
        write_run(p0, "STUDENT'S ACADEMIC CARD ", bold=True)

    if len(doc.paragraphs) > 1:
        p1 = doc.paragraphs[1]
        for r in list(p1.runs):
            p1._p.remove(r._r)
        clean_paragraph(p1)
        write_run(p1, "Faculty of Physics main department Full-time education 2022", bold=True)

    # 2. Колонтитулы
    sec = doc.sections[0]
    if sec.footer.paragraphs:
        fp = sec.footer.paragraphs[0]
        for r in list(fp.runs):
            fp._p.remove(r._r)
        write_run(fp, f"				{PLACEHOLDER_DATE}")

    # 3. Таблица 0 — Анкета условного Иванова Ивана
    t0 = doc.tables[0]
    for row in t0.rows:
        for cell in row.cells:
            if cell.paragraphs:
                clean_paragraph(cell.paragraphs[0])

    dummy_profile = {
        0: 'Ivanov Ivan',
        1: '01.01.2004',
        2: 'No contract',
        3: 'Federal budget of Russian Federation',
        4: 'Fundamental and applied Physics',
        5: '',
        7: 'Russian Federation, Moscow',
        8: 'Moscow, Leninskie Gory st., 1, 101',
        9: 'Resides',
    }
    for r_idx, val in dummy_profile.items():
        set_cell_value(t0.cell(r_idx, 1), val)
        if r_idx != 5:
            set_cell_value(t0.cell(r_idx, 2), '')
            set_cell_value(t0.cell(r_idx, 3), '')

    set_cell_value(t0.cell(6, 1), '402b')
    set_cell_value(t0.cell(6, 2), '402')
    set_cell_value(t0.cell(6, 3), '402')

    # 4. Сохраняем прототип семестровой таблицы (1 шапка + 1 строка предмета) и таблицы инспектора
    sem_table_proto = copy.deepcopy(doc.tables[1]._tbl)
    inspector_table_proto = copy.deepcopy(doc.tables[-1]._tbl)

    body = doc.element.body
    idx_after_t0 = body.index(t0._tbl) + 1
    for child in list(body[idx_after_t0:]):
        if not child.tag.endswith('sectPr'):
            body.remove(child)

    # 5. Пример блока Orders
    p_ord = doc.add_paragraph()
    clean_paragraph(p_ord, keep_with_next=True)
    write_run(p_ord, 'Orders', bold=True)

    p_ord_line = doc.add_paragraph()
    clean_paragraph(p_ord_line)
    write_run(p_ord_line, 'Accepted on the 1st course	1	25.08.2022 order N 1001-ac	20.08.2022 from 01.09.2022')

    # 6. Единственная эталонная таблица 1-го семестра (шапка + 1 строка-прототип)
    p_sp = doc.add_paragraph()
    clean_paragraph(p_sp)

    p_sem = doc.add_paragraph()
    clean_paragraph(p_sem, keep_with_next=True)
    write_run(p_sem, 'Performance of the curriculum in the 1st semester', bold=True)

    body.insert(len(body) - 1, sem_table_proto)
    sem_tbl = docx.table.Table(sem_table_proto, doc)
    while len(sem_tbl.rows) > 2:
        sem_tbl._tbl.remove(sem_tbl.rows[-1]._tr)

    for c_h in sem_tbl.rows[0].cells:
        set_cell_value(c_h, c_h.text.strip(), bold=True)

    sample_row = [
        'Module "Mathematics" Mathematical analysis',
        '6cr – 216',
        '144',
        'Exam',
        'excellent',
        '20.01.2023'
    ]
    for col_i, val in enumerate(sample_row):
        set_cell_value(sem_tbl.rows[1].cells[col_i], val)

    p_att = doc.add_paragraph()
    clean_paragraph(p_att)
    write_run(p_att, 'Attitude to the session:				', bold=True, italic=True)
    write_run(p_att, 'Passed session with “excellent” marks', bold=False, italic=True)

    # 7. Обезличивание таблицы подписи инспектора
    p_sp2 = doc.add_paragraph()
    clean_paragraph(p_sp2)

    insp_tbl = docx.table.Table(inspector_table_proto, doc)
    for row in insp_tbl.rows:
        unique_tcs = []
        for c in row.cells:
            if c._tc not in unique_tcs:
                unique_tcs.append(c._tc)
        row_txt = ' '.join(docx.table._Cell(tc, insp_tbl).text for tc in unique_tcs).lower()
        if 'inspector' in row_txt:
            rc = docx.table._Cell(unique_tcs[-1], insp_tbl)
            for extra_p in rc.paragraphs[1:]:
                rc._tc.remove(extra_p._p)
            p0 = rc.paragraphs[0]
            for r in list(p0.runs):
                p0._p.remove(r._r)
            clean_paragraph(p0, line_spacing=1.5, align=WD_ALIGN_PARAGRAPH.CENTER)
            write_run(p0, Placeholder_INSPECTOR, bold=False)
        elif 'seal' in row_txt:
            rc = docx.table._Cell(unique_tcs[-1], insp_tbl)
            for extra_p in rc.paragraphs[1:]:
                rc._tc.remove(extra_p._p)
            p0 = rc.paragraphs[0]
            for r in list(p0.runs):
                p0._p.remove(r._r)
            clean_paragraph(p0, line_spacing=1.5, align=WD_ALIGN_PARAGRAPH.CENTER)
            write_run(p0, PLACEHOLDER_DATE, bold=True)

    body.insert(len(body) - 1, inspector_table_proto)

    for line_text in [
        'I certify this is an accurate translation of the document		',
        f'{SIGNER_NAME}			',
        'International Relations Office				',
        'Faculty of Physics, MSU'
    ]:
        p_end = doc.add_paragraph()
        clean_paragraph(p_end)
        write_run(p_end, line_text)

    # 8. Очистка стандартных свойств документа (core.xml)
    cp = doc.core_properties
    cp.author = ''
    cp.last_modified_by = ''
    cp.title = ''
    cp.subject = ''
    cp.keywords = ''
    cp.comments = ''
    cp.category = ''
    cp.content_status = ''
    cp.revision = 1
    neutral_dt = datetime(2026, 1, 1, 0, 0, 0, tzinfo=timezone.utc)
    cp.created = neutral_dt
    cp.modified = neutral_dt

    doc.save(target_docx)
    scrub_docx_zip_metadata(target_docx)
    print(f"[OK] Обезличенный шаблон сохранен: {target_docx}")


def create_sample_card_xlsx(xlsx_path):
    """Создает демонстрационный файл sample_card.xlsx (лист CartSt) на условного Иванова Ивана Ивановича."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'CartSt'

    # Стили оформления демонстрационной таблицы
    font_title = Font(name='Times New Roman', size=12, bold=True, color='1F2937')
    font_sec = Font(name='Times New Roman', size=11, bold=True, color='1F2937')
    font_th = Font(name='Times New Roman', size=10, bold=True, color='FFFFFF')
    font_label = Font(name='Times New Roman', size=10, bold=True, color='374151')
    font_val = Font(name='Times New Roman', size=10, bold=False, color='111827')
    font_italic = Font(name='Times New Roman', size=10, italic=True, color='1F2937')

    fill_th = PatternFill(fill_type='solid', fgColor='334155')       # приглушенный темно-сланцевый
    fill_sec = PatternFill(fill_type='solid', fgColor='E2E8F0')      # светло-серо-голубой для секций
    fill_zebra = PatternFill(fill_type='solid', fgColor='F8FAFC')    # легкая зебра строк
    fill_profile = PatternFill(fill_type='solid', fgColor='F1F5F9')

    thin_side = Side(style='thin', color='CBD5E1')
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # Строки 1-4: Заголовок карточки (индексы строк в точности соответствуют выгрузке АИС «Студент»)
    ws.cell(3, 1, 'УЧЕБНАЯ КАРТОЧКА СТУДЕНТА').font = font_title
    ws.cell(4, 1, 'ФИЗИЧЕСКИЙ ФАКУЛЬТЕТ  основное отделение  Очная форма обучения 2022').font = font_sec

    # Строки 6-15: Анкета студента
    profile_rows = [
        (6,  ['Фамилия, имя, отчество', 'Иванов', 'Иван', 'Иванович', 'N диплома', '']),
        (7,  ['Дата рождения', '01.01.2004 г.', '', '', 'Рег. номер', '']),
        (8,  ['Договор', 'без договора', '', '', '', '']),
        (9,  ['Финансирование', 'Федеральный бюджет РФ', '', '', '', '']),
        (10, ['Специальность', 'Фундаментальная и прикладная физика', '', '', '', '']),
        (11, ['Специализация', '', '', '', 'экз.  гр.', 'стип.  гр.']),
        (12, ['Группа', '402б', '', '', '402', 402]),
        (13, ['Постоянная прописка:  регион  место', 'Российская Федерация, г. Москва', '', '', '', '']),
        (14, ['Адрес ', 'МОСКВА,ЛЕНИНСКИЕ ГОРЫ УЛ,1,101,+79990000000', '', '', '', '']),
        (15, ['Общежитие', 'проживает', '', '', '', '']),
    ]
    for r_num, vals in profile_rows:
        for c_idx, v in enumerate(vals, start=1):
            cell = ws.cell(r_num, c_idx, v)
            cell.font = font_label if c_idx == 1 else font_val
            cell.border = border_cell
            if c_idx == 1:
                cell.fill = fill_profile

    # Строки 17-24 (индексы 16-23): Скрытые шаблонные строки АИС «Студент» для проверки фильтра скрытых строк
    hidden_template_rows = [
        (17, ['Выполнение учебного плана в 1 семестре']),
        (18, ['Название  предмета', 'Всего часов', 'Аудит', 'Форма. отч.', 'Пересдачи', 'Оценка', 'Дата посл, оценки', '', 'Примечание']),
        (19, ['Отношение к сессии']),
        (20, ['Приказы']),
        (21, ['']), (22, ['']), (23, ['']), (24, ['']),
    ]
    for r_num, vals in hidden_template_rows:
        for c_idx, v in enumerate(vals, start=1):
            ws.cell(r_num, c_idx, v)
        ws.row_dimensions[r_num].hidden = True

    # Строки 25-29: Приказы (тестируют автоопределение 4-го курса и инспектора N.A. Nesterova)
    r_cur = 25
    c_ord = ws.cell(r_cur, 1, 'Приказы')
    c_ord.font = font_sec
    c_ord.fill = fill_sec
    r_cur += 1

    orders = [
        'зачислен   на 1 курс    1    25.08.2022  приказ N 1001ас   20.08.2022  с 01.09.2022',
        'переведен на следующий курс  на следующий курс    2    04.07.2023  приказ N 501-ас   30.06.2023  с 30.06.2023',
        'переведен на следующий курс  на следующий курс    3    03.07.2024  приказ N 502-ас   01.07.2024  с 01.07.2024',
        'переведен на следующий курс  на следующий курс    4    02.07.2025  приказ N 503-ас   30.06.2025  с 30.06.2025',
    ]
    for ord_txt in orders:
        cell = ws.cell(r_cur, 1, ord_txt)
        cell.font = font_val
        r_cur += 1

    r_cur += 1  # пустая строка-разделитель

    # Демонстрационные семестры (проверяют словарь модулей, courses1.txt и онлайн-перевод темы курсовой)
    semesters_demo = [
        (1, 'сдал сессию на "хорошо" и "отлично"', [
            ['Модуль "Математика" Математический анализ', '6кр - 216', '144', 'зачет', '3', 'зачтено', '21.12.2022', '', ''],
            ['Модуль "Математика" Математический анализ', '6кр - 216', '144', 'экзамен', '5', 'отлично', '16.01.2023', '', '*'],
            ['Модуль "Математика" Аналитическая геометрия', '3кр - 108', '54', 'экзамен', '5', 'отлично', '21.01.2023', '', '*'],
            ['Модуль "Общая физика" Механика', '6кр - 216', '126', 'экзамен', '4', 'хорошо', '11.01.2023', '', '*'],
            ['Модуль "Общая физика" Общий физический практикум', '3кр - 108', '72', 'зачет', '3', 'зачтено', '23.12.2022', '', ''],
            ['Модуль "Информатика" Программирование и информатика', '3кр - 108', '54', 'зачет', '3', 'зачтено', '27.12.2022', '', ''],
            ['Английский язык', '2кр - 72', '36', 'зачет', '3', 'зачтено', '24.12.2022', '', ''],
            ['Физическая культура', '2кр - 72', '72', 'зачет', '3', 'зачтено', '28.12.2022', '', '*'],
        ]),
        (2, 'сдал сессию на "отлично"', [
            ['Модуль "Теоретическая физика" Теоретическая механика', '3кр - 108', '72', 'экзамен', '5', 'отлично', '14.06.2023', '', '*'],
            ['Модуль "Общая физика" Молекулярная физика', '5кр - 180', '102', 'экзамен', '5', 'отлично', '09.06.2023', '', '*'],
            ['Модуль "Математика" Линейная алгебра', '3кр - 108', '51', 'экзамен', '5', 'отлично', '20.06.2023', '', '*'],
            ['Научное программирование на языке Python', '2кр - 72', '36', 'зачет', '3', 'зачтено', '25.05.2023', '', '*'],
            ['Техника вычисления диаграмм Фейнмана (часть 1)', '3кр - 108', '72', 'экзамен', '5', 'отлично', '24.06.2023', '', '*'],
            ['КР Лаборатория специализации: Численное моделирование квантовых спиновых цепочек', '5кр - 180', '0', 'экзамен', '5', 'отлично', '30.05.2023', '', '*'],
        ]),
    ]

    headers_row = ['Название  предмета', 'Всего часов', 'Аудит', 'Форма. отч.', 'Пересдачи', 'Оценка', 'Дата посл, оценки', '', 'Примечание']

    for sem_num, att_text, courses in semesters_demo:
        c_s = ws.cell(r_cur, 1, f'Выполнение учебного плана в {sem_num} семестре')
        c_s.font = font_sec
        c_s.fill = fill_sec
        r_cur += 1

        for c_idx, h_txt in enumerate(headers_row, start=1):
            hc = ws.cell(r_cur, c_idx, h_txt)
            hc.font = font_th
            hc.fill = fill_th
            hc.alignment = Alignment(horizontal='center' if c_idx > 1 else 'left', vertical='center')
            hc.border = border_cell
        r_cur += 1

        for idx_c, c_row in enumerate(courses):
            for c_idx, val in enumerate(c_row, start=1):
                dc = ws.cell(r_cur, c_idx, val)
                dc.font = font_val
                dc.border = border_cell
                if idx_c % 2 == 1:
                    dc.fill = fill_zebra
                if c_idx in (2, 3, 4, 5, 6, 7, 9):
                    dc.alignment = Alignment(horizontal='center', vertical='center')
            r_cur += 1

        c_att_lbl = ws.cell(r_cur, 1, 'Отношение к сессии')
        c_att_lbl.font = font_italic
        c_att_val = ws.cell(r_cur, 2, att_text)
        c_att_val.font = font_italic
        r_cur += 2

    # Секция сверх плана («Дополнительно сдано» -> Extracurricular courses)
    c_ex = ws.cell(r_cur, 1, 'Дополнительно сдано')
    c_ex.font = font_sec
    c_ex.fill = fill_sec
    r_cur += 1

    for c_idx, h_txt in enumerate(headers_row, start=1):
        hc = ws.cell(r_cur, c_idx, h_txt)
        hc.font = font_th
        hc.fill = fill_th
        hc.alignment = Alignment(horizontal='center' if c_idx > 1 else 'left', vertical='center')
        hc.border = border_cell
    r_cur += 1

    extra_course = ['КР Техника и физика ядерного эксперимента:', '1кр - 36', '18', 'зачет', '3', 'зачтено', '15.09.2025', '', '']
    for c_idx, val in enumerate(extra_course, start=1):
        dc = ws.cell(r_cur, c_idx, val)
        dc.font = font_val
        dc.border = border_cell
        if c_idx > 1:
            dc.alignment = Alignment(horizontal='center', vertical='center')

    # Настройка ширины колонок для аккуратного отображения в Excel
    col_widths = {1: 68, 2: 14, 3: 10, 4: 14, 5: 12, 6: 18, 7: 18, 8: 4, 9: 12}
    for col_idx, width in col_widths.items():
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    ws.views.sheetView[0].showGridLines = True

    # Очистка метаданных самой книги Excel
    wb.properties.creator = ''
    wb.properties.lastModifiedBy = ''
    wb.properties.title = 'Sample Student Academic Card'

    wb.save(xlsx_path)
    print(f"[OK] Демонстрационный Excel-файл создан: {xlsx_path}")


def find_source_docx(cli_arg=None):
    if cli_arg and os.path.exists(cli_arg):
        return cli_arg
    default_tpl = os.path.join(SCRIPT_DIR, 'template.docx')
    if os.path.exists(default_tpl):
        return default_tpl
    candidates = [
        os.path.join(SCRIPT_DIR, f) for f in os.listdir(SCRIPT_DIR)
        if f.lower().endswith('.docx') and not f.startswith('~$') and not f.endswith('_EN.docx')
    ]
    if candidates:
        return candidates[0]
    raise FileNotFoundError("Положите исходный .docx-образец в папку со скриптом или передайте путь аргументом.")


if __name__ == '__main__':
    src_arg = sys.argv[1] if len(sys.argv) > 1 else None
    src_docx = find_source_docx(src_arg)
    out_tpl = os.path.join(SCRIPT_DIR, 'template.docx')
    out_xlsx = os.path.join(SCRIPT_DIR, 'sample_card.xlsx')

    sanitize_template_docx(src_docx, out_tpl)
    create_sample_card_xlsx(out_xlsx)
