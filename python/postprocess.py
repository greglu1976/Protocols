#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Заменяет стили параграфов в DOCX:
1. В основном тексте: 'Обычный' -> 'Основной текст с отступом 31'
2. Подписи таблиц (содержат слово 'Таблица'): 'Обычный' -> 'СтильПодписиТаблицы'
3. В таблицах:
   - Заголовок (1-я строка): 'ДОК Таблица Текст Центр'
   - Первый столбец: 'ДОК Таблица Текст Нумерованный'
   - Остальные ячейки: 'ДОК Таблица Текст Обычный'

Использование:
    python postprocess.py путь/к/protocol_XXX.docx
"""

import sys
import os
from docx import Document

# --- НАСТРОЙКИ СТИЛЕЙ ---
TARGET_STYLE_NAME = "Основной текст с отступом 31"
CAPTION_STYLE_NAME = "ДОК Таблица Текст Без Нумерации"  # <-- Стиль для подписей типа "Таблица 1 - ..."
HEADER_STYLE_NAME = "ДОК Таблица Текст Центр"       
FIRST_COL_STYLE_NAME = "ДОК Таблица Текст Центр"        
OTHER_CELLS_STYLE_NAME = "ДОК Таблица Текст"  

def replace_style(paragraph, target_style):
    """Меняет стиль, если текущий - 'Обычный'/'Normal'."""
    if paragraph.style is None:
        return False
    
    if paragraph.style.name in ("Обычный", "Normal"):
        paragraph.style = target_style
        return True
    return False

def main():
    if len(sys.argv) < 2:
        print("Usage: python postprocess.py <docx_path>", file=sys.stderr)
        sys.exit(1)

    docx_path = sys.argv[1]

    if not os.path.isfile(docx_path):
        print(f"File not found: {docx_path}", file=sys.stderr)
        sys.exit(1)

    print(f"Post-processing: {docx_path}")
    doc = Document(docx_path)

    # Проверяем наличие всех требуемых стилей
    available_styles = {s.name for s in doc.styles}
    required_styles = [
        TARGET_STYLE_NAME, 
        CAPTION_STYLE_NAME,
        HEADER_STYLE_NAME, 
        FIRST_COL_STYLE_NAME, 
        OTHER_CELLS_STYLE_NAME
    ]
    
    missing_styles = [s for s in required_styles if s not in available_styles]
    if missing_styles:
        print(f"Error: Missing styles in document: {missing_styles}", file=sys.stderr)
        print(f"Available styles: {sorted(available_styles)}", file=sys.stderr)
        sys.exit(2)

    # Получаем объекты стилей
    style_body = doc.styles[TARGET_STYLE_NAME]
    style_caption = doc.styles[CAPTION_STYLE_NAME]
    style_header = doc.styles[HEADER_STYLE_NAME]
    style_first_col = doc.styles[FIRST_COL_STYLE_NAME]
    style_other_cells = doc.styles[OTHER_CELLS_STYLE_NAME]

    replaced = 0

    # 1. Основной текст документа
    for p in doc.paragraphs:
        # Сначала проверяем, является ли это подписью таблицы
        if "Таблица" in p.text:
            if replace_style(p, style_caption):
                replaced += 1
        else:
            # Если это не подпись, применяем основной стиль тела документа
            if replace_style(p, style_body):
                replaced += 1

    # 2. Таблицы
    for table in doc.tables:
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