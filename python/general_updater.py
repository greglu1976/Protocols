#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
general_updater.py <vars.json> <templates/general.md> <CABINET/general.md> <CABINET>

1. Читает vars.json.
2. Генерирует .xlsx для каждого устройства (ЮНИТ, ИЧМ, ПРМ/ПРД).
3. Генерирует analog.xlsx — шапка И метаданные (название, тэг) берутся
   из шаблона templates/analog.xlsx. Количество строк данных считается
   по кодам code_order / code_order2 / code_order3
   (см. calculate_analog_rows).
4. Вставляет блоки INSERT_TABLE в general.md после заголовка
   "# Основные технические данные устройств".

Шаблон templates/analog.xlsx содержит:
  - лист «Лист1» — только шапку (2 строки, с нужными объединениями,
    шрифтами и ширинами);
  - лист «Info» — название таблицы и тэг.

Строки данных добавляются программно: n_rows пустых строк, первая колонка
заполняется пробелом, иначе openpyxl не сохранит пустые строки.
"""
import json
import re
import sys
from pathlib import Path

import openpyxl


# ─── Путь к шаблону analog.xlsx ───────────────────────────────

ANALOG_TEMPLATE = Path(__file__).resolve().parent.parent / "templates" / "analog.xlsx"


# ─── Генерация xlsx ───────────────────────────────────────────

def write_table_xlsx(path, info_name, info_tag, rows):
    base, suffix = info_name

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Лист1"

    ws["A1"] = "Параметр"
    ws["B1"] = "Значение"

    r = 2
    for row in rows:
        for c, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            if isinstance(v, str) and v.startswith("="):
                cell.data_type = "s"
        r += 1

    ws_info = wb.create_sheet("Info")
    ws_info.append(["Название таблицы", base + suffix])
    ws_info.append(["Тэг", info_tag])

    wb.save(path)


def make_unit_xlsx(cabinet_dir, idx, code_order, code_unit):
    path = cabinet_dir / f"unit{idx}.xlsx"
    write_table_xlsx(
        path,
        info_name=["Технические данные ЮНИТ", f" (A{idx})"],
        info_tag=f"unit{idx}",
        rows=[
            ["Код заказа ЮНИТ", code_order],
            ["Заводской номер", ""],
            ["Завод-изготовитель", "ООО Юнител Инжиниринг"],
            ["Год выпуска", ""],
            ["Uпит, В", ""],
        ],
    )
    return path.name


def make_hmi_xlsx(cabinet_dir, unit_idx, code_hmi):
    path = cabinet_dir / f"unit{unit_idx}.1.xlsx"
    write_table_xlsx(
        path,
        info_name=["Технические данные ИЧМ", f" (A{unit_idx}.1)"],
        info_tag=f"unit{unit_idx}.1",
        rows=[
            ["Код заказа ИЧМ", code_hmi],
            ["Заводской номер", ""],
            ["Завод-изготовитель", "ООО Юнител Инжиниринг"],
            ["Год выпуска", ""],
            ["Uпит, В", ""],
        ],
    )
    return path.name


def make_prv_xlsx(cabinet_dir, code_prmprd):
    path = cabinet_dir / "prv.xlsx"
    write_table_xlsx(
        path,
        info_name=["Технические данные ВЧ ПРМ/ПРД", ""],
        info_tag="prv",
        rows=[
            ["Код заказа ВЧ ПРМ/ПРД", code_prmprd],
            ["Заводской номер", ""],
            ["Завод-изготовитель", "ООО Юнител Инжиниринг"],
            ["Год выпуска", ""],
            ["Uпит, В", ""],
        ],
    )
    return path.name


# ─── Аналоговые каналы: подсчёт строк и xlsx ──────────────────

# Кириллические буквы, визуально совпадающие с латинскими.
_CYR_TO_LAT = {
    '\u0410': 'A',  # А
    '\u0412': 'B',  # В
    '\u0421': 'C',  # С
    '\u0415': 'E',  # Е
    '\u041d': 'H',  # Н
    '\u041a': 'K',  # К
    '\u041c': 'M',  # М
    '\u041e': 'O',  # О
    '\u0420': 'P',  # Р
    '\u0422': 'T',  # Т
    '\u0425': 'X',  # Х
}


def _char_value(ch: str) -> int:
    """Значение символа: цифра — как есть, латинская A=10 … Z=35.
    Кириллические буквы-двойники дают то же значение, что и латинские.
    Прочее — 0."""
    if ch.isdigit():
        return int(ch)
    up = ch.upper()
    if up in _CYR_TO_LAT:
        up = _CYR_TO_LAT[up]
    if 'A' <= up <= 'Z':
        return 10 + (ord(up) - ord('A'))
    return 0


def calculate_analog_rows(code_orders) -> int:
    """
    Считает суммарное количество строк для таблицы аналоговых каналов.
    Для каждого кода:
      - разбиваем по дефису (обычный / en / em / прочие);
      - находим фрагменты, начинающиеся с 'M1' / 'М1' (любой регистр);
      - берём часть после 'M1' до первой точки;
      - суммируем значения символов (A=10, B=11, C=12, цифры как есть).
    """
    total = 0
    for code in code_orders:
        if not code:
            continue
        for part in re.split(r'[-–—\u2010\u2011\u2012\u2013\u2014\u2015]', code):
            m = re.match(r'^[MmМм]1([^.]*)', part)
            if not m:
                continue
            tail = m.group(1)
            total += sum(_char_value(ch) for ch in tail)
    return total


def make_analog_xlsx(cabinet_dir, n_rows, template_xlsx=None):
    """
    Копирует шаблон templates/analog.xlsx целиком (шапка + лист Info)
    и добавляет n_rows пустых строк данных на активный лист.

    Ничего в шапке и в Info не пересобирается — всё берётся из шаблона.
    Первая колонка каждой новой строки заполняется пробелом, иначе
    openpyxl не сохранит пустую строку.
    """
    if template_xlsx is None:
        template_xlsx = ANALOG_TEMPLATE

    template_xlsx = Path(template_xlsx)
    if not template_xlsx.is_file():
        raise FileNotFoundError(
            f"Шаблон analog.xlsx не найден: {template_xlsx}. "
            f"Положите файл с шапкой и листом Info в templates/."
        )

    wb = openpyxl.load_workbook(template_xlsx)
    ws = wb.active

    # Первая строка после шапки. В шаблоне только шапка (2 строки),
    # max_row вернёт 2 — значит данные начнутся с 3-й строки.
    start_row = ws.max_row + 1
    for r in range(start_row, start_row + n_rows):
        ws.cell(row=r, column=1, value=" ")

    path = cabinet_dir / "analog.xlsx"
    wb.save(path)
    return path.name


# ─── Блоки INSERT_TABLE ───────────────────────────────────────

def make_insert_block(xlsx_name, sheet="Лист1"):
    return (
        f"<!-- INSERT_TABLE:file={xlsx_name} -->\n"
        f"<!-- END_INSERT_TABLE -->"
    )


# ─── main ─────────────────────────────────────────────────────

def main():
    if len(sys.argv) != 5:
        print(
            "Usage: general_updater.py <vars.json> <template.md> <out.md> <cabinet_dir>",
            file=sys.stderr,
        )
        return 2

    vars_json   = Path(sys.argv[1])
    template_md = Path(sys.argv[2])
    out_md      = Path(sys.argv[3])
    cabinet_dir = Path(sys.argv[4])

    vars_data = json.loads(vars_json.read_text(encoding="utf-8"))

    units_codes = [v for k in ("code_order", "code_order2", "code_order3")
                   if (v := vars_data.get(k)) is not None]
    hmis  = [v for k in ("code_hmi", "code_hmi2", "code_hmi3")
             if (v := vars_data.get(k)) is not None]
    prv   = vars_data.get("code_prmprd")

    inserts = []

    # ── ЮНИТ + ИЧМ ──
    for i, code_order in enumerate(units_codes, start=1):
        make_unit_xlsx(cabinet_dir, i, code_order, vars_data.get("code_unit", ""))
        inserts.append(make_insert_block(f"unit{i}.xlsx"))

        if i - 1 < len(hmis):
            make_hmi_xlsx(cabinet_dir, i, hmis[i - 1])
            inserts.append(make_insert_block(f"unit{i}.1.xlsx"))

    # ── ВЧ ПРМ/ПРД ──
    if prv:
        make_prv_xlsx(cabinet_dir, prv)
        inserts.append(make_insert_block("prv.xlsx"))

    # ── Аналоговые каналы ──
    n_analog_rows = calculate_analog_rows(units_codes)
    make_analog_xlsx(cabinet_dir, n_analog_rows)
    print(f"[general_updater] analog.xlsx: {n_analog_rows} row(s)")

    # ── Вставка блоков INSERT_TABLE после маркера ──
    content = template_md.read_text(encoding="utf-8")

    marker = "# Основные технические данные устройств"
    idx = content.find(marker)
    if idx == -1:
        print("[general_updater] ERROR: '# Основные технические данные устройств' not found in template", file=sys.stderr)
        return 1

    end_of_line = content.find("\n", idx)
    if end_of_line == -1:
        end_of_line = len(content)

    insert_text = "\n\n".join(inserts)
    content = (
        content[:end_of_line + 1]
        + "\n"
        + insert_text
        + "\n\n"
        + content[end_of_line + 1:]
    )

    out_md.parent.mkdir(parents=True, exist_ok=True)
    out_md.write_text(content, encoding="utf-8")

    print(f"[general_updater] generated {len(inserts)} table(s) -> {out_md}")
    return 0


if __name__ == "__main__":
    sys.exit(main())