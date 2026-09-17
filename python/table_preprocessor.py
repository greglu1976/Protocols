#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Препроцессор таблиц для протоколов.
Заменяет теги <!-- INSERT_TABLE: ... --> на ASCII-таблицы из Excel.
Поддерживает относительные пути и автоматическое обновление данных.
"""
import sys
import os
import re
import math
import openpyxl
from pathlib import Path

# Регулярка для поиска тегов
TAG_PATTERN = re.compile(
    r'<!--\s*(?:INSERT_TABLE|TABLE_PLACEHOLDER):\s*'
    r'file=(?P<file>[^,\s]+)'
    r'(?:,\s*sheet=(?P<sheet>[^,]+?))?'
    r'\s*-->'
)


def excel_to_perfect_grid(file_path, sheet_name=0):
    """Генерирует ASCII-сетку таблицы из Excel с подписью для pandoc-crossref."""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.worksheets[sheet_name] if isinstance(sheet_name, int) else wb[sheet_name]

    max_row = ws.max_row
    max_col = ws.max_column
    merged_ranges = ws.merged_cells.ranges

    def get_merged_info(r, c):
        for rng in merged_ranges:
            if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
                is_master = (r == rng.min_row and c == rng.min_col)
                return is_master, rng.min_row, rng.min_col, rng.max_row, rng.max_col
        return True, r, c, r, c

    # 1. Сбор данных и расчет ширины колонок
    col_widths = [3] * max_col
    raw_data = {}
    spanned_cells = []

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            is_master, min_r, min_c, max_r, max_c = get_merged_info(r, c)
            val = ""
            if is_master:
                val = str(ws.cell(row=r, column=c).value or '').replace('\n', ' ').strip()

            raw_data[(r, c)] = {
                'val': val, 'is_master': is_master,
                'min_r': min_r, 'min_c': min_c, 'max_r': max_r, 'max_c': max_c
            }

            if is_master:
                if min_c == max_c:
                    col_widths[c - 1] = max(col_widths[c - 1], len(val))
                else:
                    spanned_cells.append((min_c, max_c, len(val)))

    # Корректируем ширину под colspan
    spanned_cells.sort(key=lambda x: (x[1] - x[0]))
    for min_c, max_c, val_len in spanned_cells:
        current_total_width = sum(col_widths[min_c - 1:max_c]) + (max_c - min_c)
        if val_len > current_total_width:
            needed_extra = val_len - current_total_width
            num_cols = max_c - min_c + 1
            extra_per_col = math.ceil(needed_extra / num_cols)
            for c_idx in range(min_c, max_c + 1):
                col_widths[c_idx - 1] += extra_per_col

    row_heights = [1] * max_row

    # 2. Вычисление символьных координат X и Y
    col_starts = [0] * (max_col + 1)
    current_x = 0
    for i, w in enumerate(col_widths):
        col_starts[i] = current_x
        current_x += w + 1
    col_starts[max_col] = current_x

    row_starts = [0] * (max_row + 1)
    current_y = 0
    for i, h in enumerate(row_heights):
        row_starts[i] = current_y
        current_y += h + 1
    row_starts[max_row] = current_y

    total_width = col_starts[-1] + 1
    total_height = row_starts[-1] + 1

    # 3. Инициализация холста
    canvas = [[' ' for _ in range(total_width)] for _ in range(total_height)]

    # Заполняем базовую сетку
    for r in range(max_row + 1):
        y = row_starts[r]
        for x in range(total_width - 1):
            canvas[y][x] = '-'
    for c in range(max_col + 1):
        x = col_starts[c]
        for y in range(total_height - 1):
            canvas[y][x] = '|'
    for r in range(max_row + 1):
        for c in range(max_col + 1):
            canvas[row_starts[r]][col_starts[c]] = '+'

    # 4. Затирание внутренних границ для объединенных диапазонов
    unique_ranges = set()
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            _, min_r, min_c, max_r, max_c = get_merged_info(r, c)
            if min_r != max_r or min_c != max_c:
                unique_ranges.add((min_r, min_c, max_r, max_c))

    for min_r, min_c, max_r, max_c in unique_ranges:
        y1 = row_starts[min_r - 1]
        y2 = row_starts[max_r]
        x1 = col_starts[min_c - 1]
        x2 = col_starts[max_c]

        for y_inner in range(y1 + 1, y2):
            for x_inner in range(x1 + 1, x2):
                canvas[y_inner][x_inner] = ' '

        for y_inner in range(y1 + 1, y2):
            if y_inner not in row_starts:
                continue
            for x_inner in range(x1 + 1, x2):
                if canvas[y_inner][x_inner] in ['-', '+']:
                    canvas[y_inner][x_inner] = ' '

        for x_inner in range(x1 + 1, x2):
            if x_inner not in col_starts:
                continue
            for y_inner in range(y1 + 1, y2):
                if canvas[y_inner][x_inner] in ['|', '+']:
                    canvas[y_inner][x_inner] = ' '

    # 5. Печать текста
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            info = raw_data[(r, c)]
            if info['is_master']:
                y1 = row_starts[info['min_r'] - 1]
                x1 = col_starts[info['min_c'] - 1]
                x2 = col_starts[info['max_c']]

                inner_width = (x2 - x1) - 1
                padded_val = info['val'].ljust(inner_width)[:inner_width]

                for char_idx, char in enumerate(padded_val):
                    canvas[y1 + 1][x1 + 1 + char_idx] = char

    # 6. Двойная линия разделителя заголовка
    header_end_row = 2
    if header_end_row <= max_row:
        header_y = row_starts[header_end_row]
        for x in range(total_width - 1):
            if canvas[header_y][x] == '-':
                canvas[header_y][x] = '='

    # 7. Парсинг имени листа и формирование подписи
    sheet_title = ws.title
    match = re.match(r'^(.*?)\(([^)]+)\)\s*$', sheet_title)

    if match:
        visible_name = match.group(1).strip()
        table_id = match.group(2).strip()
        caption_line = f": {visible_name} {{#tbl:{table_id}}}"
    else:
        safe_id = sheet_title.replace(' ', '_')
        caption_line = f": {sheet_title} {{#tbl:{safe_id}}}"

    grid_output = '\n'.join(''.join(row).rstrip() for row in canvas)
    return f"{grid_output}\n\n{caption_line}"


def replace_tag(match, md_file_dir):
    """Заменяет найденный тег на ASCII-таблицу."""
    params = match.groupdict()
    file_rel = params['file']
    sheet_name = params.get('sheet', 0)

    # Резолвим путь относительно директории MD-файла
    file_abs = os.path.normpath(os.path.join(md_file_dir, file_rel))

    if not os.path.isfile(file_abs):
        print(f"[PREPROCESS] ERROR: File not found: {file_abs}", file=sys.stderr)
        return f"\n[ERROR: Table file '{file_rel}' not found]\n"

    try:
        table_ascii = excel_to_perfect_grid(file_abs, sheet_name=sheet_name)
        print(f"[PREPROCESS] Inserted table from '{file_rel}' (sheet='{sheet_name}')")
        return "\n" + table_ascii + "\n"
    except Exception as e:
        print(f"[PREPROCESS] ERROR processing '{file_rel}': {e}", file=sys.stderr)
        return f"\n[ERROR: Failed to process table '{file_rel}']\n"


def process_file(md_path):
    """Обрабатывает один Markdown-файл."""
    content = Path(md_path).read_text(encoding='utf-8')
    md_file_dir = os.path.dirname(os.path.abspath(md_path))

    matches_count = len(TAG_PATTERN.findall(content))
    if matches_count == 0:
        return content

    print(f"[PREPROCESS] Found {matches_count} table tag(s) in {md_path}")
    new_content = TAG_PATTERN.sub(lambda m: replace_tag(m, md_file_dir), content)
    return new_content


def main():
    if len(sys.argv) < 2:
        print("Usage: python table_preprocessor.py file1.md [file2.md ...]", file=sys.stderr)
        sys.exit(1)

    processed_count = 0
    for md_file in sys.argv[1:]:
        if not os.path.isfile(md_file):
            print(f"[PREPROCESS] Warning: File not found: {md_file}", file=sys.stderr)
            continue

        result = process_file(md_file)
        Path(md_file).write_text(result, encoding='utf-8')
        processed_count += 1

    print(f"[PREPROCESS] Done. Processed {processed_count} file(s).")


if __name__ == "__main__":
    main()