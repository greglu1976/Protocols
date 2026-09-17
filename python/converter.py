import openpyxl
import math

def excel_to_perfect_grid(file_path, sheet_name=0):
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
                    col_widths[c-1] = max(col_widths[c-1], len(val))
                else:
                    spanned_cells.append((min_c, max_c, len(val)))

    # Корректируем ширину под colspan
    spanned_cells.sort(key=lambda x: (x[1] - x[0]))
    for min_c, max_c, val_len in spanned_cells:
        current_total_width = sum(col_widths[min_c-1:max_c]) + (max_c - min_c)
        if val_len > current_total_width:
            needed_extra = val_len - current_total_width
            num_cols = max_c - min_c + 1
            extra_per_col = math.ceil(needed_extra / num_cols)
            for c_idx in range(min_c, max_c + 1):
                col_widths[c_idx-1] += extra_per_col

    # Каждая строка Excel = 1 текстовая строка на холсте
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
        for x in range(total_width - 1): canvas[y][x] = '-'
    for c in range(max_col + 1):
        x = col_starts[c]
        for y in range(total_height - 1): canvas[y][x] = '|'
    for r in range(max_row + 1):
        for c in range(max_col + 1): canvas[row_starts[r]][col_starts[c]] = '+'

    # 4. Затирание внутренних границ для объединенных диапазонов
    # Чтобы не стереть лишнего, собираем уникальные диапазоны слияния
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

        # Затираем ВСЕ внутренние разделители строго ВНУТРИ объединенной области
        # Границы (y1, y2, x1, x2) не трогаем, стираем только то, что между ними
        for y_inner in range(y1 + 1, y2):
            for x_inner in range(x1 + 1, x2):
                canvas[y_inner][x_inner] = ' '
        
        # Стираем внутренние перегородки '|' на стыках строк внутри слияния
        for y_inner in range(y1 + 1, y2):
            if y_inner not in row_starts:
                continue
            for x_inner in range(x1 + 1, x2):
                if canvas[y_inner][x_inner] in ['-', '+']:
                    canvas[y_inner][x_inner] = ' '

        # Стираем внутренние перегородки '-' на стыках колонок внутри слияния
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

    # 6. Отрисовка двойной линии разделителя заголовка (после ВТОРОЙ строки Excel, т.к. заголовок двухстрочный)
    # Определяем, где кончается самый глубокий элемент заголовка. Обычно это строка 2.
    header_end_row = 2 
    if header_end_row <= max_row:
        header_y = row_starts[header_end_row]
        for x in range(total_width - 1):
            if canvas[header_y][x] == '-':
                canvas[header_y][x] = '='

    return '\n'.join(''.join(row).rstrip() for row in canvas)

# --- Запуск ---
print(excel_to_perfect_grid('table.xlsx'))
