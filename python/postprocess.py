#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Заменяет стили параграфов в DOCX и задаёт ширины столбцов таблиц.
Плюс правит word/numbering.xml: маркеры списков → тире,
абзац по левому краю, первая строка — с отступом 1 см,
таб-стоп — на 1.5 см.
Также вставляет таблицу подписей на место маркера @@SIGNATURE_TABLE@@.
"""
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

import sys
import os
import re
import shutil
import zipfile
from docx import Document

# --- МАРКЕР ДЛЯ ВСТАВКИ ТАБЛИЦЫ ПОДПИСЕЙ ---
SIGNATURE_TABLE_MARKER = "%%SIGNATURE_TABLE%%"

# --- НАСТРОЙКИ СТИЛЕЙ ---
TARGET_STYLE_NAME = "Основной текст с отступом 31"
CAPTION_STYLE_NAME = "ДОК Таблица Текст Без Нумерации"
HEADER_STYLE_NAME = "ДОК Таблица Текст Центр"
FIRST_COL_STYLE_NAME = "ДОК Таблица Текст Центр"
OTHER_CELLS_STYLE_NAME = "ДОК Таблица Текст Центр"
IMAGE_STYLE_NAME = "ДОК Базовый с маленьким интервалом"

# --- МАРКЕР СПИСКА ---
NEW_BULLET_MARKER = "–"

# --- ОТСТУП СПИСКА (twips, 1 см = 567) ---
LIST_LEFT_TWIPS       = 0       # весь абзац — по левому краю
LIST_FIRSTLINE_TWIPS  = 567     # первая строка (с маркером) — 1 см
LIST_TAB_TWIPS        = 850     # таб-стоп для текста после маркера — 1.5 см

# --- КАРТА ТИПОВ ТАБЛИЦ ---
TABLE_TYPE_KEYWORDS = {
    "вход для проверки": "NewTable",
    "параметр": "NewTable2",
    "№ п/п": "PerechenTable"
}

# --- КАРТА ПОДТИПОВ ТАБЛИЦ ---
# Ключ — имя подтипа (как в TABLE_WIDTHS_PCT),
# Значение — подстрока, которую ищем в шапке таблицы (в нижнем регистре).
# Порядок важен: первый найденный подтип побеждает.
SUBTYPE_KEYWORDS = {
    "I2гарм": "i2гарм",   # сначала проверяем более специфичный
    "Iторм":  "iторм",
    # сюда добавляйте новые подтипы — без правок в detect_table_subtype
}

# --- КАРТА ШИРИН ---
# Ключ: (type_name, n_cols)                — базовый вариант
#       (type_name, n_cols, subtype)       — уточнённый вариант (приоритетный)
TABLE_WIDTHS_PCT = {
    # ── двухколоночные "параметр" ──
    ("NewTable2", 2):  [30, 70],

    # ── NewTable без подтипа (старые записи) ──
    ("NewTable", 6):  [16, 32, 16, 12, 12, 12],
    ("NewTable", 7):  [16, 32, 16, 9, 9, 9, 9],
    ("NewTable", 10): [15, 23, 15, 8, 6, 6, 6, 7, 7, 7],

    # ── NewTable, 8 столбцов — два подтипа ──
    # fallback (если подтип не определён)
    ("NewTable", 8):           [16, 33, 16, 7, 7, 7, 7, 7],
    # подтип "Iторм" — таблица характеристики ДТЗт
    ("NewTable", 8, "Iторм"):  [16, 26, 13, 9, 9, 9, 9, 9],
    ("NewTable", 8, "I2гарм"):  [15, 23, 12, 10, 10, 10, 10, 10],

    # ── PerechenTable ──
    ("PerechenTable", 7): [5, 30, 10, 10, 10, 11, 24],
}

# --- ШИРИНЫ СТОЛБЦОВ ТАБЛИЦЫ ПОДПИСЕЙ (см) ---
SIGNATURE_TABLE_WIDTHS_CM = [4.5, 4.0, 5.5, 3.0]


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
    """
    Универсальный определитель подтипа.
    Перебирает SUBTYPE_KEYWORDS в порядке объявления
    и возвращает первый подтип, чья подстрока встречается в шапке.
    Если ничего не найдено — возвращает None.
    """
    if not table.rows:
        return None

    header_cells = [c.text.strip().lower() for c in table.rows[0].cells]
    header_joined = " ".join(header_cells)

    for subtype_name, keyword in SUBTYPE_KEYWORDS.items():
        if keyword in header_joined:
            return subtype_name
    return None


def detect_table_type(table):
    """
    Возвращает (type_name, n_cols, subtype).
    subtype = None, если подтип не определён.
    """
    if not table.rows:
        return None, 0, None

    n_cols = len(table.columns)
    header_text = table.rows[0].cells[0].text.strip().lower()

    for keyword, type_name in TABLE_TYPE_KEYWORDS.items():
        if keyword in header_text:
            subtype = detect_table_subtype(table)
            return type_name, n_cols, subtype

    return None, n_cols, None


def set_cell_text(cell, text, bold=False, align="left", size=None):
    """Записывает текст в ячейку с нужным форматированием."""
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


def set_bottom_border(cell):
    """Рисует нижнюю границу ячейки, не затирая остальные границы."""
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
    """Возвращает первый абзац, содержащий указанный текст, или None."""
    for p in doc.paragraphs:
        if text in p.text:
            return p
    return None


def clear_table_borders(table):
    """Убирает все границы у таблицы целиком."""
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
    """Снимает все границы у одной ячейки."""
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
    """Снимает границы у всех ячеек таблицы."""
    for row in table.rows:
        for cell in row.cells:
            clear_cell_borders(cell)


def apply_normal_style_to_cells(table, doc):
    """
    Применяет стиль "Обычный1" ко всем абзацам во всех ячейках таблицы.
    """
    normal_style = None
    for name in ("Обычный1", "Normal1"):
        try:
            normal_style = doc.styles[name]
            break
        except KeyError:
            continue

    if normal_style is None:
        print("  [warn] Стиль 'Обычный'/'Normal' не найден — пропуск",
              file=sys.stderr)
        return

    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                try:
                    p.style = normal_style
                except Exception as e:
                    print(f"  [warn] Не удалось применить стиль: {e}",
                          file=sys.stderr)


def build_signature_table(doc):
    # 7 строк:
    #   0 — пустая (вверху)
    #   1 — "Проверку произвели:" + линии/подписи
    #   2 — пустая (нижняя половина merge)
    #   3 — "Протокол проверил:" + линии/подписи
    #   4 — пустая (нижняя половина merge)
    #   5 — пустая (между "Протокол проверил:" и "М.П.")
    #   6 — "М.П." + текст-предупреждение
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

    # ── merge в первом столбце ──
    table.cell(1, 0).merge(table.cell(2, 0))
    table.cell(3, 0).merge(table.cell(4, 0))

    # "М.П." и текст-предупреждение — в последней строке
    table.cell(6, 1).merge(table.cell(6, 3))

    # ── Текст в левом столбце ──
    set_cell_text(table.cell(1, 0), "Проверку произвели:")
    set_cell_text(table.cell(3, 0), "Протокол проверил:")
    set_cell_text(table.cell(6, 0), "М.П.", bold=True, align="center")

    # ── Линии + подписи в ОДНОЙ ячейке (строки 1 и 3) ──
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

    # ── Текст-предупреждение ──
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

    # ── Снять границы ──
    clear_all_cell_borders(table)
    clear_table_borders(table)

    # ── Стиль шрифта ячеек → "Обычный" ──
    apply_normal_style_to_cells(table, doc)

    return table


def set_cell_text_two_lines(cell, line_text, caption_text,
                            line_size=11, caption_size=9):
    """
    Кладёт в ячейку два параграфа:
      1) линию (подчёркивания)
      2) подпись под ней
    Оба по центру.
    """
    while len(cell.paragraphs) > 1:
        p = cell.paragraphs[-1]._p
        p.getparent().remove(p)

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


def insert_table_at_marker(doc, marker, table_builder):
    """
    Находит абзац с маркером, удаляет его и вставляет на его место
    таблицу, построенную table_builder(doc).
    """
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
        "Обычный", "Normal", "Основной текст", "Body Text"
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
    """
    Правки в word/numbering.xml:
      - у bullet-уровней (lvlText = одиночный символ) символ → new_marker,
        и убирается <w:rFonts>, чтобы маркер рисовался обычным шрифтом;
      - у ВСЕХ уровней: <w:tabs> с таб-стопом на tab_twips
        и <w:ind w:left w:firstLine>.
    """
    tmp = docx_path + ".tmp"
    shutil.move(docx_path, tmp)

    replaced_count = 0
    indent_count = 0

    LVL_TEXT_RE = re.compile(r'<w:lvlText\s+w:val="([^"]*)"\s*/>')

    def process_lvl(match):
        nonlocal replaced_count, indent_count
        block = match.group(0)

        # ── 1. bullet или numbered? ──
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

        # ── 2. Отступы + таб-стоп ──
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

    print(f"  Bullet markers replaced: {replaced_count} → '{new_marker}'")
    print(f"  List indents set: {indent_count} level(s), "
          f"left={left_twips}, firstLine={firstline_twips}, "
          f"tab={tab_twips} twips")
    return replaced_count, indent_count


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

    available_styles = {s.name for s in doc.styles}
    required_styles = [
        TARGET_STYLE_NAME, CAPTION_STYLE_NAME,
        HEADER_STYLE_NAME, FIRST_COL_STYLE_NAME, OTHER_CELLS_STYLE_NAME,
        IMAGE_STYLE_NAME,
    ]
    missing_styles = [s for s in required_styles if s not in available_styles]
    if missing_styles:
        print(f"Error: Missing styles in document: {missing_styles}",
              file=sys.stderr)
        print(f"Available styles: {sorted(available_styles)}", file=sys.stderr)
        sys.exit(2)

    style_body        = doc.styles[TARGET_STYLE_NAME]
    style_caption     = doc.styles[CAPTION_STYLE_NAME]
    style_header      = doc.styles[HEADER_STYLE_NAME]
    style_first_col   = doc.styles[FIRST_COL_STYLE_NAME]
    style_other_cells = doc.styles[OTHER_CELLS_STYLE_NAME]
    style_image       = doc.styles[IMAGE_STYLE_NAME]

    replaced = 0

    # 1. Основной текст
    for p in doc.paragraphs:
        if paragraph_has_image(p):
            if replace_style(p, style_image, force=True):
                replaced += 1
            print(f"  Image paragraph style set: {p.style.name}")
        elif "Таблица" in p.text:
            if replace_style(p, style_caption):
                replaced += 1
        else:
            if replace_style(p, style_body):
                replaced += 1

    # 2. Таблицы
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

        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                current_target_style = style_other_cells
                if row_idx == 0:
                    current_target_style = style_header
                elif cell_idx == 0:
                    current_target_style = style_first_col

                for p in cell.paragraphs:
                    if replace_style(p, current_target_style):
                        replaced += 1

    # 3. Вставка таблицы подписей на место маркера
    insert_table_at_marker(
        doc,
        SIGNATURE_TABLE_MARKER,
        build_signature_table,
    )

    print(f"Replaced paragraphs: {replaced}")
    doc.save(docx_path)

    # 4. numbering.xml: маркеры + отступы + таб
    fix_lists_in_numbering(docx_path)

    print("Saved:", docx_path)
    print("Done.")


if __name__ == "__main__":
    main()