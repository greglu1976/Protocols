#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
general_updater.py <vars.json> <templates/general.md> <CABINET/general.md> <CABINET>

1. Читает vars.json.
2. Генерирует .xlsx для каждого устройства (ЮНИТ, ИЧМ, ПРМ/ПРД).
3. Вставляет блоки INSERT_TABLE в general.md после заголовка "# Основные технические данные устройств".
"""
import json
import sys
from pathlib import Path

import openpyxl


# ─── Генерация xlsx ───────────────────────────────────────────

def write_table_xlsx(path, info_name, info_tag, rows):
    base, suffix = info_name

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Лист1"

    # Строка 1 — объединённый заголовок (без суффикса)
    ws.merge_cells("A1:B1")
    ws["A1"] = base
    ws["A1"].alignment = openpyxl.styles.Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )

    # Строка 2 — шапка
    ws["A2"] = "Параметр"
    ws["B2"] = "Значение"

    # Строки 3+ — данные
    r = 3
    for row in rows:
        for c, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            if isinstance(v, str) and v.startswith("="):
                cell.data_type = "s"
        r += 1

    # Лист Info — метаданные (с суффиксом)
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

    units = [v for k in ("code_order", "code_order2", "code_order3")
             if (v := vars_data.get(k)) is not None]
    hmis  = [v for k in ("code_hmi", "code_hmi2", "code_hmi3")
             if (v := vars_data.get(k)) is not None]
    prv   = vars_data.get("code_prmprd")

    inserts = []

    for i, code_order in enumerate(units, start=1):
        make_unit_xlsx(cabinet_dir, i, code_order, vars_data.get("code_unit", ""))
        inserts.append(make_insert_block(f"unit{i}.xlsx"))

        if i - 1 < len(hmis):
            make_hmi_xlsx(cabinet_dir, i, hmis[i - 1])
            inserts.append(make_insert_block(f"unit{i}.1.xlsx"))

    if prv:
        make_prv_xlsx(cabinet_dir, prv)
        inserts.append(make_insert_block("prv.xlsx"))

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