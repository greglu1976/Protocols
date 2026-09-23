#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Заменяет стили параграфов в DOCX и задаёт ширины столбцов таблиц
в зависимости от типа таблицы.

1. В основном тексте: 'Обычный' -> 'Основной текст с отступом 31'
2. Подписи таблиц (содержат слово 'Таблица'): 'Обычный' -> 'СтильПодписиТаблицы'
3. В таблицах:
   - Заголовок (1-я строка): 'ДОК Таблица Текст Центр'
   - Первый столбец: 'ДОК Таблица Текст Центр'
   - Остальные ячейки: 'ДОК Таблица Текст Центр'
4. Ширины столбцов таблиц задаются по типу (см. TABLE_WIDTHS_PCT).

Тип таблицы определяется по ключевому слову в заголовке первого столбца:
  - "Ток"    -> AnalogueTable
  - "Сигнал" -> DiscreteTable

Использование:
    python postprocess.py путь/к/protocol_XXX.docx
"""

from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import sys
import os
from docx import Document

# --- НАСТРОЙКИ СТИЛЕЙ ---
TARGET_STYLE_NAME = "Основной текст с отступом 31"
CAPTION_STYLE_NAME = "ДОК Таблица Текст Без Нумерации"
HEADER_STYLE_NAME = "ДОК Таблица Текст Центр"
FIRST_COL_STYLE_NAME = "ДОК Таблица Текст Центр"
OTHER_CELLS_STYLE_NAME = "ДОК Таблица Текст Центр"

# --- КАРТА ТИПОВ ТАБЛИЦ ---
# Ключ — подстрока в заголовке первого столбца (в нижнем регистре),
# значение — имя типа таблицы.
TABLE_TYPE_KEYWORDS = {
    "вход для проверки": "NewTable",      
}

# --- КАРТА ШИРИН ---
# Ключ: (имя_типа, число_столбцов)
# Значение: список процентов, сумма = 100.
#
# ЗАМЕНИТЕ значения на нужные вам. Ниже — примерные заглушки,
# чтобы скрипт запускался «из коробки».
TABLE_WIDTHS_PCT = {
    ("NewTable", 6):  [16, 32, 16, 12, 12, 12],    
    ("NewTable", 7):  [16, 32, 16, 9, 9, 9, 9],
    ("NewTable", 8):  [16, 33, 16, 7, 7, 7, 7, 7],    
    ("NewTable", 10):  [15, 23, 15, 8, 6, 6, 6, 7, 7, 7],     
}

# ============================================================
#  Работа с XML таблицы
# ============================================================

def set_table_layout_fixed(table):
    """Принудительно фиксированный layout, иначе Word пересчитает ширины."""
    tblPr = table._tbl.tblPr
    layout = tblPr.find(qn('w:tblLayout'))
    if layout is None:
        layout = OxmlElement('w:tblLayout')
        tblPr.append(layout)
    layout.set(qn('w:type'), 'fixed')

def set_table_width_percent(table, percent=100):
    """Устанавливает ширину таблицы в процентах от доступной ширины."""
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.append(tblW)
    tblW.set(qn('w:w'), str(percent * 50))
    tblW.set(qn('w:type'), 'pct')

def set_table_grid(table, widths_pct):
    """
    Прописывает ширины столбцов в tblGrid и в каждой ячейке.
    widths_pct: список процентов, сумма = 100.
    """
    total = sum(widths_pct)
    if total != 100:
        raise ValueError(f"Сумма ширин должна быть 100, а не {total}")

    tbl = table._tbl
    tblPr = tbl.tblPr

    # 1. Ширина таблицы = 100%
    set_table_width_percent(table, 100)

    # 2. Фиксированный layout
    set_table_layout_fixed(table)

    # 3. Переписываем tblGrid
    old_grid = tbl.find(qn('w:tblGrid'))
    if old_grid is not None:
        tbl.remove(old_grid)
    grid = OxmlElement('w:tblGrid')
    for pct in widths_pct:
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(50 * pct))  # 5000 twips = 100%
        grid.append(gc)
    tblPr.addnext(grid)

    # 4. Прописываем ширину каждой ячейки
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

def detect_table_type(table):
    """Возвращает (type_name, n_cols) или (None, n_cols)."""
    if not table.rows:
        return None, 0
    n_cols = len(table.columns)
    header_text = table.rows[0].cells[0].text.strip().lower()
    for keyword, type_name in TABLE_TYPE_KEYWORDS.items():
        if keyword in header_text:
            return type_name, n_cols
    return None, n_cols

# ============================================================
#  Работа со стилями
# ============================================================

def replace_style(paragraph, target_style):
    """Меняет стиль, если текущий — 'Обычный'/'Normal'/'Основной текст'/'Body Text'."""
    if paragraph.style is None:
        return False
    if paragraph.style.name in ("Обычный", "Normal", "Основной текст", "Body Text"):
        paragraph.style = target_style
        return True
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

    # Проверяем наличие всех требуемых стилей
    available_styles = {s.name for s in doc.styles}
    required_styles = [
        TARGET_STYLE_NAME,
        CAPTION_STYLE_NAME,
        HEADER_STYLE_NAME,
        FIRST_COL_STYLE_NAME,
        OTHER_CELLS_STYLE_NAME,
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

    replaced = 0

    # 1. Основной текст документа
    for p in doc.paragraphs:
        if "Таблица" in p.text:
            if replace_style(p, style_caption):
                replaced += 1
        else:
            if replace_style(p, style_body):
                replaced += 1

    # 2. Таблицы
    for table in doc.tables:
        type_name, n_cols = detect_table_type(table)

        if type_name is not None:
            widths = TABLE_WIDTHS_PCT.get((type_name, n_cols))
            if widths is not None:
                set_table_grid(table, widths)
                print(f"  {type_name} ({n_cols} cols): widths = {widths}%")
            else:
                print(
                    f"  [warn] Нет ширин для {type_name} с {n_cols} столбцами — пропуск",
                    file=sys.stderr,
                )
                set_table_width_percent(table, 100)
        else:
            # Fallback: просто растянуть на 100%
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

    print(f"Replaced paragraphs: {replaced}")
    doc.save(docx_path)
    print("Saved:", docx_path)
    print("Done.")


if __name__ == "__main__":
    main()