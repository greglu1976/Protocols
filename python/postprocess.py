#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Заменяет стили параграфов в DOCX и задаёт ширины столбцов таблиц.
Плюс правит word/numbering.xml: маркеры списков -> тире,
абзац по левому краю, первая строка - с отступом 1 см,
таб-стоп - на 1.5 см.
Также вставляет таблицу подписей на место маркера %%SIGNATURE_TABLE%%.
И приклеивает в конец документа лист ревизий (revsheet.docx),
путь к которому задаётся через переменную окружения REVSHEET.
"""
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

import copy
import sys
import os
import re
import shutil
import zipfile
import json
from docx import Document

# --- слияние docx (необязательная зависимость) ---
try:
    from docxcompose.composer import Composer
    _HAS_DOCXCOMPOSE = True
except ImportError:
    _HAS_DOCXCOMPOSE = False

# --- МАРКЕР ДЛЯ ВСТАВКИ ТАБЛИЦЫ ПОДПИСЕЙ ---
SIGNATURE_TABLE_MARKER = "%%SIGNATURE_TABLE%%"

# --- ИМЯ ПЕРЕМЕННОЙ ОКРУЖЕНИЯ С ПУТЁМ К ЛИСТУ РЕВИЗИЙ ---
REVSHEET_ENV = "REVSHEET"
VARS_JSON_ENV = "VARS_JSON"

# --- НАСТРОЙКИ СТИЛЕЙ ---
TARGET_STYLE_NAME = "Основной текст с отступом 31"
CAPTION_STYLE_NAME = "ДОК Таблица Текст Без Нумерации"
HEADER_STYLE_NAME = "ДОК Таблица Текст Центр"
FIRST_COL_STYLE_NAME = "ДОК Таблица Текст Центр"
OTHER_CELLS_STYLE_NAME = "ДОК Таблица Текст Центр"
IMAGE_STYLE_NAME = "ДОК Базовый с маленьким интервалом"

# --- ЕДИНЫЙ РАЗМЕР ШРИФТА В ТАБЛИЦАХ (pt) ---
TABLE_FONT_PT = 10.0

# --- ИНТЕРВАЛ ПЕРЕД АБЗАЦАМИ В ЯЧЕЙКАХ ТАБЛИЦ (pt) ---
TABLE_SPACING_BEFORE_PT = 0.0

# --- МАРКЕР СПИСКА ---
NEW_BULLET_MARKER = "–"

# --- ОТСТУП СПИСКА (twips, 1 см = 567) ---
LIST_LEFT_TWIPS = 0
LIST_FIRSTLINE_TWIPS = 567
LIST_TAB_TWIPS = 850

# --- КАРТА ТИПОВ ТАБЛИЦ ---
TABLE_TYPE_KEYWORDS = {
    "вход для проверки": "NewTable",
    "параметр": "NewTable2",
    "№ п/п": "PerechenTable",
    "наименование задействованного": "DiscretTable",
    "измерительный канал": "AnalogTable",
}

# --- КАРТА ПОДТИПОВ ТАБЛИЦ ---
SUBTYPE_KEYWORDS = {
    "I2гарм": "i2гарм",
    "Iторм": "iторм",
}

# --- КАРТА ШИРИН ---
TABLE_WIDTHS_PCT = {
    ("NewTable2", 2): [30, 70],
    ("NewTable", 6): [16, 32, 16, 12, 12, 12],
    ("NewTable", 7): [16, 32, 16, 9, 9, 9, 9],
    ("NewTable", 10): [15, 23, 15, 8, 6, 6, 6, 7, 7, 7],
    ("NewTable", 8): [16, 33, 16, 7, 7, 7, 7, 7],
    ("NewTable", 8, "Iторм"): [16, 26, 13, 9, 9, 9, 9, 9],
    ("NewTable", 8, "I2гарм"): [15, 23, 12, 10, 10, 10, 10, 10],
    ("PerechenTable", 7): [5, 30, 10, 10, 10, 11, 24],
    ("DiscretTable", 4): [38, 12, 38, 12],
    ("AnalogTable", 7): [30, 12, 8, 12, 8, 15, 15],
}

# --- ШИРИНЫ СТОЛБЦОВ ТАБЛИЦЫ ПОДПИСЕЙ (см) ---
SIGNATURE_TABLE_WIDTHS_CM = [4.5, 4.0, 5.5, 3.0]

# Сопоставление названий столбцов и строк в таблице изоляции
INSUL_COL_TO_ROW = {
    "ТТ": "Цепи переменного тока (ТТ)",
    "ТН": "Цепи переменного напряжения \"звезды\" (ТН)",
    "ШОН": "Цепи отбора напряжения (ШОН)",
    "EC": "Цепи постоянного опер. Тока (EC)",
    "Вых. цепи": "Выходные цепи",
    "EH": "Цепи сигнализации (EH)",
    "РАС": "Цепи РАС",
    "Осв.": "Цепи освещения шкафа"
}

# ============================================================
#  Правки титульного листа: разрывы строк и формат дат
# ============================================================
_MONTHS = {
    "января": 1, "февраля": 2, "марта": 3, "апреля": 4,
    "мая": 5, "июня": 6, "июля": 7, "августа": 8,
    "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12,
}

_DATE_RE = re.compile(
    r"(\d{1,2})[ \u00A0]+"
    r"(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)"
    r"[ \u00A0]+(\d{4})(?:[ \u00A0]*(?:г\.|года))?"
)

_ISO_DATE_RE = re.compile(r"\b(\d{4})-(\d{2})-(\d{2})\b")
_TBL_ID_RE = re.compile(r"\{#tbl:[^}]*\}")


def _iter_all_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def _strip_tbl_artifacts(doc):
    removed = 0
    for p in _iter_all_paragraphs(doc):
        runs = p.runs
        if not runs:
            continue
        full = "".join(r.text for r in runs)
        if "{#tbl:" not in full:
            continue
        new = _TBL_ID_RE.sub("", full)
        if new != full:
            runs[0].text = new
            for r in runs[1:]:
                r.text = ""
            removed += 1
    return removed


def _fix_literal_newlines(paragraph):
    fixed = 0
    for run in list(paragraph.runs):
        if "\\n" not in run.text:
            continue
        parts = run.text.split("\\n")
        run.text = parts[0].replace("\\r", "")
        anchor = run._r
        for part in parts[1:]:
            new_r = copy.deepcopy(run._r)
            for t in new_r.findall(qn('w:t')):
                new_r.remove(t)
            br = OxmlElement('w:br')
            new_r.append(br)
            t = OxmlElement('w:t')
            t.set(qn('xml:space'), 'preserve')
            t.text = part.replace("\\r", "")
            new_r.append(t)
            anchor.addnext(new_r)
            anchor = new_r
        fixed += 1
    return fixed


def _fix_dates(paragraph):
    n = 0
    for run in paragraph.runs:
        def repl_m(m):
            nonlocal n
            n += 1
            return "{:02d}.{:02d}.{}".format(int(m.group(1)), _MONTHS[m.group(2)], m.group(3))
        new = _DATE_RE.sub(repl_m, run.text)
        if new != run.text:
            run.text = new

        def repl_iso(m):
            nonlocal n
            n += 1
            return "{:02d}.{:02d}.{}".format(int(m.group(3)), int(m.group(2)), m.group(1))
        new = _ISO_DATE_RE.sub(repl_iso, run.text)
        if new != run.text:
            run.text = new
    return n


def _collapse_double_spaces(paragraph):
    runs = paragraph.runs
    if not runs:
        return 0
    full = "".join(r.text for r in runs)
    if not re.search(r"[ \u00A0]{2,}(?=\d)", full):
        return 0
    new = re.sub(r"[ \u00A0]{2,}(?=\d)", " ", full)
    runs[0].text = new
    for r in runs[1:]:
        r.text = ""
    return 1


# ============================================================
#  Таблицы
# ============================================================
def set_table_layout_fixed(table):
    tblPr = table._tbl.tblPr
    layout = tblPr.find(qn('w:tblLayout'))
    if layout is None:
        layout = OxmlElement('w:tblLayout')
        tblPr.append(layout)
    layout.set(qn('w:type'), 'fixed')


def set_table_width_percent(table, percent=100):
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.append(tblW)
    tblW.set(qn('w:w'), str(percent * 50))
    tblW.set(qn('w:type'), 'pct')


def set_table_grid(table, widths_pct):
    total = sum(widths_pct)
    if total != 100:
        raise ValueError(f"Сумма ширин должна быть 100, а не {total}")
    tbl = table._tbl
    tblPr = tbl.tblPr
    set_table_width_percent(table, 100)
    set_table_layout_fixed(table)
    old_grid = tbl.find(qn('w:tblGrid'))
    if old_grid is not None:
        tbl.remove(old_grid)
    grid = OxmlElement('w:tblGrid')
    for pct in widths_pct:
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(50 * pct))
        grid.append(gc)
    tblPr.addnext(grid)
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            if idx >= len(widths_pct):
                continue
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = OxmlElement('w:tcW')
                tcPr.append(tcW)
            tcW.set(qn('w:w'), str(50 * widths_pct[idx]))
            tcW.set(qn('w:type'), 'pct')


def detect_table_subtype(table):
    if not table.rows:
        return None
    header_cells = [c.text.strip().lower() for c in table.rows[0].cells]
    header_joined = " ".join(header_cells)
    for subtype_name, keyword in SUBTYPE_KEYWORDS.items():
        if keyword in header_joined:
            return subtype_name
    return None


def detect_table_type(table):
    if not table.rows:
        return None, 0, None
    n_cols = len(table.columns)
    header_text = table.rows[0].cells[0].text.strip().lower()
    for keyword, type_name in TABLE_TYPE_KEYWORDS.items():
        if keyword in header_text:
            subtype = detect_table_subtype(table)
            return type_name, n_cols, subtype
    return None, n_cols, None


def normalize_table_sizes(doc, table_pt=TABLE_FONT_PT):
    fixed = 0
    for table in doc.tables:
        for r, row in enumerate(table.rows):
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(table_pt)
                        if r > 0:
                            run.font.bold = None
                        fixed += 1
    return fixed


def normalize_table_spacing(doc, before_pt=TABLE_SPACING_BEFORE_PT):
    n = 0
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(before_pt)
                    n += 1
    return n


def _drop_empty_paragraphs(cell):
    for p in list(cell.paragraphs)[1:]:
        if not p.text.strip():
            p._p.getparent().remove(p._p)


def merge_analog_cells(table):
    if len(table.columns) != 7 or len(table.rows) < 3:
        return
    table.cell(0, 0).merge(table.cell(1, 0))
    for col in (1, 3, 5):
        table.cell(0, col).merge(table.cell(0, col + 1))
    for col in (0, 1, 3, 5):
        _drop_empty_paragraphs(table.cell(0, col))
    for r in range(2, len(table.rows)):
        first = table.cell(r, 0).text.strip()
        if first.startswith("«"):
            continue
        if first and all(not table.cell(r, c).text.strip() for c in range(1, 7)):
            base = table.cell(r, 0)
            for c in range(1, 7):
                cell = table.cell(r, c)
                if cell._tc is not base._tc:
                    base = base.merge(cell)
            _drop_empty_paragraphs(base)
            for p in base.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _shade_cell(cell, fill_color="d9d9d9"):
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag == qn('w:shd'):
            tcPr.remove(child)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    tcPr.append(shd)


def _delete_column(table, col_idx):
    for row in table.rows:
        cell = row.cells[col_idx]
        row._tr.remove(cell._tc)
    tbl = table._tbl
    grid = tbl.find(qn('w:tblGrid'))
    if grid is not None:
        grid_cols = grid.findall(qn('w:gridCol'))
        if col_idx < len(grid_cols):
            grid.remove(grid_cols[col_idx])


def _delete_row(table, row_idx):
    tr = table.rows[row_idx]._tr
    table._tbl.remove(tr)


def process_insul_table(doc, allowed_circuits=None):
    """
    Находит таблицу изоляции, удаляет лишние строки/столбцы
    и закрашивает нижний треугольник.
    """
    for table in doc.tables:
        if len(table.rows) > 1 and len(table.columns) > 1:
            header_row = [cell.text.strip() for cell in table.rows[0].cells]
            if "ТТ" in header_row and "ТН" in header_row and "ШОН" in header_row:
                
                if allowed_circuits is not None:
                    cols_to_delete = []
                    rows_to_delete = []
                    
                    for idx, col_name in enumerate(header_row):
                        if idx > 0 and col_name not in allowed_circuits:
                            cols_to_delete.append(idx)
                            
                    for r_idx in range(1, len(table.rows)):
                        row_text = table.cell(r_idx, 0).text.strip()
                        for col_idx in cols_to_delete:
                            col_name = header_row[col_idx]
                            if INSUL_COL_TO_ROW.get(col_name) == row_text:
                                rows_to_delete.append(r_idx)
                                break
                    
                    for col_idx in sorted(cols_to_delete, reverse=True):
                        _delete_column(table, col_idx)
                        
                    for r_idx in sorted(rows_to_delete, reverse=True):
                        _delete_row(table, r_idx)
                        
                    print(f"  Таблица изоляции отфильтрована. Оставлены цепи: {allowed_circuits}")

                # Заливка
                num_cols = len(table.columns)
                for r_idx in range(1, len(table.rows)):
                    for c_idx in range(1, min(r_idx + 1, num_cols)):
                        cell = table.cell(r_idx, c_idx)
                        _shade_cell(cell)
                
                print(f"  Заливка для таблицы изоляции успешно применена.")
                break


def set_cell_text(cell, text, bold=False, align="left", size=None):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text)
    run.bold = bold
    if size:
        run.font.size = Pt(size)
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def set_cell_text_two_lines(cell, line_text, caption_text,
                            line_size=11, caption_size=9):
    while len(cell.paragraphs) > 1:
        p = cell.paragraphs[-1]._p
        p.getparent().remove(p._p)
    p1 = cell.paragraphs[0]
    for r in list(p1.runs):
        r._r.getparent().remove(r._r)
    run1 = p1.add_run(line_text)
    run1.font.size = Pt(line_size)
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2 = cell.add_paragraph()
    run2 = p2.add_run(caption_text)
    run2.font.size = Pt(caption_size)
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER


def set_bottom_border(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn('w:tcBorders'))
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tcPr.append(borders)
    old_bottom = borders.find(qn('w:bottom'))
    if old_bottom is not None:
        borders.remove(old_bottom)
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:color'), '000000')
    borders.append(bottom)


def find_paragraph_by_text(doc, text):
    for p in doc.paragraphs:
        if text in p.text:
            return p
    return None


def clear_table_borders(table):
    tblPr = table._tbl.tblPr
    old = tblPr.find(qn('w:tblBorders'))
    if old is not None:
        tblPr.remove(old)
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'none')
        el.set(qn('w:sz'), '0')
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), 'auto')
        borders.append(el)
    tblPr.append(borders)


def clear_cell_borders(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    old = tcPr.find(qn('w:tcBorders'))
    if old is not None:
        tcPr.remove(old)
    borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'none')
        el.set(qn('w:sz'), '0')
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), 'auto')
        borders.append(el)
    tcPr.append(borders)


def clear_all_cell_borders(table):
    for row in table.rows:
        for cell in row.cells:
            clear_cell_borders(cell)


def apply_normal_style_to_cells(table, doc):
    normal_style = None
    for name in ("Обычный1", "Normal1"):
        try:
            normal_style = doc.styles[name]
            break
        except KeyError:
            continue
    if normal_style is None:
        print("  [warn] Стиль 'Обычный'/'Normal' не найден — пропуск", file=sys.stderr)
        return
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                try:
                    p.style = normal_style
                except Exception as e:
                    print(f"  [warn] Не удалось применить стиль: {e}", file=sys.stderr)


def build_signature_table(doc):
    table = doc.add_table(rows=7, cols=4)
    try:
        table.style = "Table Grid"
    except KeyError:
        pass
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            cell.width = Cm(SIGNATURE_TABLE_WIDTHS_CM[idx])
    table.cell(1, 0).merge(table.cell(2, 0))
    table.cell(3, 0).merge(table.cell(4, 0))
    table.cell(6, 1).merge(table.cell(6, 3))
    set_cell_text(table.cell(1, 0), "Проверку произвели:")
    set_cell_text(table.cell(3, 0), "Протокол проверил:")
    set_cell_text(table.cell(6, 0), "М.П.", bold=True, align="center")
    SIGNS = {
        1: ("_" * 20, "(подпись)"),
        2: ("_" * 25, "(расшифровка подписи)"),
        3: ("_" * 13, "(дата)"),
    }
    for row_idx in (1, 3):
        for col, (line, caption) in SIGNS.items():
            set_cell_text_two_lines(
                table.cell(row_idx, col),
                line, caption,
                line_size=11,
                caption_size=9,
            )
    merged_cell = table.cell(6, 1)
    merged_cell.text = ""
    lines = [
        "Частичная или полная перепечатка и размножение только "
        "с разрешения испытательной лаборатории.",
        "Исправления не допускаются.",
        "Протокол распространяется только на элементы электроустановки "
        "или оборудования, подвергнутые измерениям (проверке).",
    ]
    for i, line in enumerate(lines):
        p = merged_cell.paragraphs[0] if i == 0 else merged_cell.add_paragraph()
        run = p.add_run(line)
        run.font.size = Pt(11)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    clear_all_cell_borders(table)
    clear_table_borders(table)
    apply_normal_style_to_cells(table, doc)
    return table


def insert_table_at_marker(doc, marker, table_builder):
    p = find_paragraph_by_text(doc, marker)
    if p is None:
        print(f"  [warn] Маркер '{marker}' не найден", file=sys.stderr)
        return False
    parent = p._p.getparent()
    index = list(parent).index(p._p)
    parent.remove(p._p)
    table = table_builder(doc)
    parent.remove(table._tbl)
    parent.insert(index, table._tbl)
    parent.insert(index + 1, OxmlElement('w:p'))
    print(f"  Таблица вставлена на место '{marker}'")
    return True


# ============================================================
#  Стили
# ============================================================

def replace_style(paragraph, target_style, force=False):
    if paragraph.style is None:
        return False
    if force or paragraph.style.name in (
        "Обычный", "Normal", "Основной текст", "Body Text",
        "Compact", "First Paragraph",
        "Table Caption", "Image Caption",
    ):
        paragraph.style = target_style
        return True
    return False


def paragraph_has_image(paragraph):
    p = paragraph._p
    return bool(p.findall('.//' + qn('w:drawing'))) or \
        bool(p.findall('.//' + qn('w:pict')))


# ============================================================
#  numbering.xml: маркеры + отступы + таб
# ============================================================

def fix_lists_in_numbering(docx_path,
                           new_marker=NEW_BULLET_MARKER,
                           left_twips=LIST_LEFT_TWIPS,
                           firstline_twips=LIST_FIRSTLINE_TWIPS,
                           tab_twips=LIST_TAB_TWIPS):
    tmp = docx_path + ".tmp"
    shutil.move(docx_path, tmp)
    replaced_count = 0
    indent_count = 0
    LVL_TEXT_RE = re.compile(r'<w:lvlText\s+w:val="([^"]*)"\s*/>')

    def process_lvl(match):
        nonlocal replaced_count, indent_count
        block = match.group(0)
        mtext = LVL_TEXT_RE.search(block)
        if mtext:
            val = mtext.group(1)
            if val == "" or len(val) == 1:
                block = re.sub(
                    r'(<w:lvlText\s+w:val=")[^"]*("\s*/>)',
                    lambda mm: mm.group(1) + new_marker + mm.group(2),
                    block, count=1,
                )
                block = re.sub(r'<w:rFonts[^/]*/>', '', block)
                replaced_count += 1
        block = re.sub(r'<w:ind[^/]*/>', '', block)
        block = re.sub(r'<w:tabs>.*?</w:tabs>', '', block, flags=re.DOTALL)
        new_inner = (
            f'<w:tabs><w:tab w:val="left" w:pos="{tab_twips}"/></w:tabs>'
            f'<w:ind w:left="{left_twips}" w:firstLine="{firstline_twips}"/>'
        )
        ppr_match = re.search(r'<w:pPr[^>]*>', block)
        if ppr_match:
            at = ppr_match.end()
            block = block[:at] + new_inner + block[at:]
        else:
            block = re.sub(r'</w:lvl>',
                           f'<w:pPr>{new_inner}</w:pPr></w:lvl>',
                           block)
        indent_count += 1
        return block

    try:
        with zipfile.ZipFile(tmp, "r") as zin, \
                zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/numbering.xml":
                    text = data.decode("utf-8")
                    text = re.sub(
                        r'<w:lvl\b.*?</w:lvl>',
                        process_lvl,
                        text,
                        flags=re.DOTALL,
                    )
                    data = text.encode("utf-8")
                zout.writestr(item, data)
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass

    print(f"  Bullet markers replaced: {replaced_count} -> '{new_marker}'")
    print(f"  List indents set: {indent_count} level(s), "
          f"left={left_twips}, firstLine={firstline_twips}, tab={tab_twips} twips")
    return replaced_count, indent_count


# ============================================================
#  Колонтитулы: формат дат в word/header*.xml и word/footer*.xml
# ============================================================

_HF_NAME_RE = re.compile(r"word/(?:header|footer)\d*\.xml$")


def fix_headers_footers(docx_path):
    tmp = docx_path + ".tmp2"
    shutil.move(docx_path, tmp)
    total = 0
    try:
        with zipfile.ZipFile(tmp, "r") as zin, \
                zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if _HF_NAME_RE.search(item.filename):
                    text = data.decode("utf-8")
                    n = len(_ISO_DATE_RE.findall(text)) + len(_DATE_RE.findall(text))
                    if n:
                        text = _ISO_DATE_RE.sub(
                            lambda m: "{:02d}.{:02d}.{}".format(
                                int(m.group(3)), int(m.group(2)), m.group(1)),
                            text,
                        )
                        text = _DATE_RE.sub(
                            lambda m: "{:02d}.{:02d}.{}".format(
                                int(m.group(1)), _MONTHS[m.group(2)], m.group(3)),
                            text,
                        )
                        data = text.encode("utf-8")
                    if n:
                        print(f"  {item.filename}: dates reformatted: {n}")
                    total += n
                zout.writestr(item, data)
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass
    print(f"  Headers/footers dates reformatted: {total}")
    return total


# ============================================================
#  Приклеивание revsheet.docx
# ============================================================

def _append_docx_raw(main_path, extra_path):
    main = Document(main_path)
    extra = Document(extra_path)
    main_body = main.element.body
    extra_body = extra.element.body
    main_sect = main_body.find(qn('w:sectPr'))
    for child in list(extra_body):
        if child.tag == qn('w:sectPr'):
            continue
        main_body.append(child)
    if main_sect is not None:
        main_body.remove(main_sect)
        main_body.append(main_sect)
    main.save(main_path)


def append_docx(main_path, extra_path):
    if not os.path.isfile(extra_path):
        print(f"  [warn] revsheet не найден: {extra_path} — пропуск", file=sys.stderr)
        return False
    if _HAS_DOCXCOMPOSE:
        base = Document(main_path)
        composer = Composer(base)
        composer.append(Document(extra_path))
        composer.save(main_path)
        print(f"  revsheet appended via docxcompose: {extra_path}")
        return True
    print("  [warn] docxcompose не установлен — используется fallback", file=sys.stderr)
    try:
        _append_docx_raw(main_path, extra_path)
        print(f"  revsheet appended (raw): {extra_path}")
        return True
    except Exception as e:
        print(f"  [error] не удалось приклеить revsheet: {e}", file=sys.stderr)
        return False


# ============================================================
#  main
# ============================================================

def main():
    if len(sys.argv) < 2:
        print("Usage: python postprocess.py <docx_path>", file=sys.stderr)
        sys.exit(1)

    docx_path = sys.argv[1]

    if not os.path.isfile(docx_path):
        print(f"File not found: {docx_path}", file=sys.stderr)
        sys.exit(1)

    if os.environ.get("TEST") == "1":
        print(f"[TEST MODE] Skipping post-processing for {docx_path}")
        return

    print(f"Post-processing: {docx_path}")
    doc = Document(docx_path)

    # 0. Правки титульного листа
    n_breaks = 0
    n_dates = 0
    for p in _iter_all_paragraphs(doc):
        n_breaks += _fix_literal_newlines(p)
        n_dates += _fix_dates(p)
    n_tbl_ids = _strip_tbl_artifacts(doc)
    print(f"  Line-breaks inserted: {n_breaks}, dates reformatted: {n_dates}, "
          f"tbl-id artifacts removed: {n_tbl_ids}")

    # 1. Нормализация размеров шрифта и интервала
    n_fixed = normalize_table_sizes(doc)
    print(f"  Normalized font sizes in tables: {n_fixed} run(s) -> {TABLE_FONT_PT} pt")
    n_spaced = normalize_table_spacing(doc)
    print(f"  Table spacing before set to {TABLE_SPACING_BEFORE_PT} pt: "
          f"{n_spaced} paragraph(s)")

    available_styles = {s.name for s in doc.styles}
    required_styles = [
        TARGET_STYLE_NAME, CAPTION_STYLE_NAME,
        HEADER_STYLE_NAME, FIRST_COL_STYLE_NAME, OTHER_CELLS_STYLE_NAME,
        IMAGE_STYLE_NAME,
    ]
    missing_styles = [s for s in required_styles if s not in available_styles]
    if missing_styles:
        print(f"Error: Missing styles in document: {missing_styles}", file=sys.stderr)
        print(f"Available styles: {sorted(available_styles)}", file=sys.stderr)
        sys.exit(2)

    style_body = doc.styles[TARGET_STYLE_NAME]
    style_caption = doc.styles[CAPTION_STYLE_NAME]
    style_header = doc.styles[HEADER_STYLE_NAME]
    style_first_col = doc.styles[FIRST_COL_STYLE_NAME]
    style_other_cells = doc.styles[OTHER_CELLS_STYLE_NAME]
    style_image = doc.styles[IMAGE_STYLE_NAME]

    replaced = 0

    # 2. Основной текст
    for p in doc.paragraphs:
        if paragraph_has_image(p):
            if replace_style(p, style_image, force=True):
                replaced += 1
            print(f"  Image paragraph style set: {p.style.name}")
        elif "Таблица" in p.text or p.style.name in ("Table Caption", "Image Caption"):
            if replace_style(p, style_caption, force=True):
                replaced += 1
        else:
            if replace_style(p, style_body):
                replaced += 1

    # 3. Схлопывание задвоенных пробелов
    n_spaces = 0
    for p in doc.paragraphs:
        n_spaces += _collapse_double_spaces(p)
    print(f"  Double spaces collapsed: {n_spaces} paragraph(s)")

    # 4. Таблицы
    for table in doc.tables:
        type_name, n_cols, subtype = detect_table_type(table)
        if type_name is not None:
            widths = None
            if subtype is not None:
                widths = TABLE_WIDTHS_PCT.get((type_name, n_cols, subtype))
            if widths is None:
                widths = TABLE_WIDTHS_PCT.get((type_name, n_cols))
            if widths is not None:
                set_table_grid(table, widths)
                sub = f", subtype={subtype}" if subtype else ""
                print(f"  {type_name} ({n_cols} cols{sub}): widths = {widths}%")
            else:
                print(f"  [warn] Нет ширин для {type_name} "
                      f"с {n_cols} столбцами — пропуск", file=sys.stderr)
                set_table_width_percent(table, 100)
        else:
            set_table_width_percent(table, 100)
        if type_name == "AnalogTable":
            merge_analog_cells(table)
        header_rows = 2 if type_name == "AnalogTable" else 1
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                current_target_style = style_other_cells
                if row_idx < header_rows:
                    current_target_style = style_header
                elif cell_idx == 0:
                    current_target_style = style_first_col
                for p in cell.paragraphs:
                    if replace_style(p, current_target_style, force=True):
                        replaced += 1

    # 4.1 ЗАЛИВКА И ФИЛЬТРАЦИЯ ТАБЛИЦЫ ИЗОЛЯЦИИ
    allowed_circuits = None
    vars_json_path = os.environ.get(VARS_JSON_ENV)
    
    if vars_json_path and os.path.isfile(vars_json_path):
        try:
            with open(vars_json_path, "r", encoding="utf-8") as f:
                vars_data = json.load(f)
            
            allowed_circuits = []
            # Ключи в vars.json теперь полностью совпадают с заголовками в таблице
            table_headers = ["ТТ", "ТН", "ШОН", "EC", "Вых. цепи", "EH", "РАС", "Осв."]
            
            for col_name in table_headers:
                val = vars_data.get(col_name)
                # Если стоит плюс (или 1, или True), добавляем в разрешенные
                if val in ("+", 1, True):
                    allowed_circuits.append(col_name)
            
            print(f"  Разрешенные цепи для таблицы изоляции: {allowed_circuits}")
            
        except Exception as e:
            print(f"  [warn] Ошибка чтения vars.json: {e}", file=sys.stderr)
            allowed_circuits = None

    process_insul_table(doc, allowed_circuits)

    # 5. Вставка таблицы подписей на место маркера
    insert_table_at_marker(doc, SIGNATURE_TABLE_MARKER, build_signature_table)

    # 5.1 Интервал перед и в таблице подписей
    n_sig_spaced = normalize_table_spacing(doc)
    print(f"  Table spacing before set to {TABLE_SPACING_BEFORE_PT} pt "
          f"(incl. signature table): {n_sig_spaced} paragraph(s)")

    print(f"Replaced paragraphs: {replaced}")
    doc.save(docx_path)

    # 6. numbering.xml
    fix_lists_in_numbering(docx_path)

    # 7. Приклеить лист ревизий
    revsheet = os.environ.get(REVSHEET_ENV)
    if revsheet:
        if not os.path.isabs(revsheet) and os.sep not in revsheet \
                and "/" not in revsheet:
            cabinet_dir = os.environ.get("CABINET_DIR")
            if cabinet_dir:
                revsheet = os.path.join(cabinet_dir, revsheet)
        append_docx(docx_path, revsheet)
    else:
        print(f"  [info] {REVSHEET_ENV} не задан — revsheet не добавляется")

    # 8. Колонтитулы
    fix_headers_footers(docx_path)

    print("Saved:", docx_path)
    print("Done.")


if __name__ == "__main__":
    main()