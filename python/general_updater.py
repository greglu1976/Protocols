#!/usr/bin/env python
# -*- coding: utf-8 -*-
import json
import os
import re
import sys
import tempfile
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from table_generator import generate_table_ascii

REPEAT_BY_FILE = {
    "unit.xlsx": ("code_order", "code_order2", "code_order3"),
    "hmi.xlsx": ("code_hmi", "code_hmi2", "code_hmi3"),
}
GATE_BY_FILE = {"prv.xlsx": "code_prmprd"}
DATA_SHEET = "Лист1"
PLACEHOLDER_RE = re.compile(r"\{\{\s*([A-Za-z0-9_]+)\s*\}\}")

DISCRET_JSON_DIR = Path(__file__).resolve().parent.parent / "discret"
ANALOG_JSON_DIR = Path(__file__).resolve().parent.parent / "analog"

GENERATED_DISCRET_IDS = []
GENERATED_ANALOG_IDS = []

DISCRET_HEADERS = (
    "Наименование задействованного входа", "U ср / U в, В",
    "Наименование задействованного входа", "U ср / U в, В",
)
_ANALOG_H1 = ("Измерительный канал", "Подаваемая величина", "",
              "Измеренная величина", "", "Максимальная погрешность, %", "")
_ANALOG_H2 = ("", "Величина, А(В)", "Фаза, °", "Величина, А(В)",
              "Фаза, °", "Амплитудная", "Угловая")
_ANALOG_MIN = (20, 16, 9, 16, 9, 12, 9)


def _nonempty(v):
    return v is not None and (not isinstance(v, str) or bool(v.strip()))


def _sanitize_id(raw):
    """Pandoc принимает в ID подписи только буквы, цифры, дефисы
    и подчёркивания — точка (unit1.1) рендерится литералом и даёт
    артефакт в DOCX. Поэтому точка и прочее -> подчёркивание."""
    s = re.sub(r"[^0-9A-Za-zА-Яа-яЁё_]+", "_", str(raw))
    s = s.strip("_") or "table"
    return s


def _caption_line(caption, crossref_id):
    # Срезаем все {#tbl:...}, которые могли попасть в текст подписи
    # из генератора или Info-листа, и добавляем ровно один чистый ID.
    caption = re.sub(r"\s*\{#tbl:[^}]*\}", "", caption).strip()
    return f": {caption} {{#tbl:{_sanitize_id(crossref_id)}}}"


def _substitute_in_xlsx(template_path, ctx):
    wb = openpyxl.load_workbook(template_path)

    def repl(m):
        val = ctx.get(m.group(1))
        return "" if val is None else str(val)

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str):
                    continue
                new = PLACEHOLDER_RE.sub(repl, v)
                if new != v:
                    cell.value = new
    return wb


def _pick_data_sheet(wb):
    if DATA_SHEET in wb.sheetnames:
        return wb[DATA_SHEET]
    for s in wb.worksheets:
        if s.title.lower() != "info":
            return s
    return wb.worksheets[0]


def _get_info_caption(wb):
    if "Info" not in wb.sheetnames:
        return None
    for row in wb["Info"].iter_rows(min_row=1, max_col=2):
        k = str(row[0].value).strip().lower() if row[0].value else ""
        if k == "название таблицы":
            v = row[1].value
            if v is not None and str(v).strip():
                return str(v).strip()
    return None


def _save_and_render(wb, ws):
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(tmp_fd)
    try:
        wb.save(tmp_path)
        return generate_table_ascii(tmp_path, sheet_name=ws.title)
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass


def _split_trailing_captions(md):
    """Отрезает хвостовые строки ': ...' и извлекает из них ID —
    если генератор уже вставил {#tbl:...} в подпись, он сохраняется
    как есть, дублей не появляется."""
    lines = md.rstrip().split("\n")
    caps = []
    found_ids = []
    while lines and lines[-1].lstrip().startswith(":"):
        cap_line = lines.pop().lstrip()[1:].strip()
        m = re.search(r"\{#tbl:([^}]*)\}", cap_line)
        if m:
            found_ids.append(m.group(1).strip())
            cap_line = re.sub(r"\s*\{#tbl:[^}]*\}", "", cap_line).strip()
        caps.append(cap_line)
    return "\n".join(lines).rstrip(), caps, found_ids


def _slugify(text):
    slug = re.sub(r"[^0-9A-Za-zА-Яа-яЁё]+", "_", text).strip("_").lower()
    return slug or "table"


def _grid_md(headers, rows, min_widths=()):
    """Grid-таблица pandoc. headers — строки шапки (последняя отделяется
    '='), rows — тело с многострочными ячейками 'text \\\nперенос'."""
    widths = [0] * len(headers[0])
    for h in headers:
        for i, t in enumerate(h):
            widths[i] = max(widths[i], len(t))
    for i, mw in enumerate(min_widths):
        widths[i] = max(widths[i], mw)
    for row in rows:
        for i, cell in enumerate(row):
            for ln in cell.split("\n"):
                widths[i] = max(widths[i], len(ln))

    def border(ch):
        return "+" + "+".join(ch * (w + 2) for w in widths) + "+"

    def line(cells):
        n = max(len(c.split("\n")) for c in cells)
        out = []
        for k in range(n):
            parts = []
            for i, c in enumerate(cells):
                ls = c.split("\n")
                v = ls[k] if k < len(ls) else ""
                parts.append(" " + v.ljust(widths[i]) + " ")
            out.append("|" + "|".join(parts) + "|")
        return "\n".join(out)

    md = []
    for h in headers:
        md.append(border("-"))
        md.append(line(list(h)))
    md.append(border("="))
    for r in rows:
        md.append(line(r))
        md.append(border("-"))
    return "\n".join(md)


def render_analog_from_json(vars_data):
    cabinet_id = str(vars_data.get("id", ""))
    json_path = None
    if ANALOG_JSON_DIR.is_dir():
        for p in ANALOG_JSON_DIR.glob("*.json"):
            if cabinet_id in p.name:
                json_path = p
                break
    if not json_path or not json_path.is_file():
        print(f"[general_updater] WARNING: JSON аналоговых {cabinet_id} не найден в {ANALOG_JSON_DIR}", file=sys.stderr)
        return ""
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f).get("data", {})
    except Exception as e:
        print(f"[general_updater] ERROR {json_path}: {e}", file=sys.stderr)
        return ""

    tables = []
    for device_name, device_data in data.items():
        if not isinstance(device_data, dict) or not device_data:
            continue
        groups = []
        for group_name, channels in device_data.items():
            if not isinstance(channels, dict):
                continue
            rows = []
            for channel_name, terminals in channels.items():
                term_str = ", ".join(str(t) for t in terminals) if isinstance(terminals, list) else str(terminals)
                rows.append(f"«{channel_name}» \\\n{term_str}")
            if rows:
                groups.append((group_name, rows))
        if not groups:
            continue
        body = []
        for gname, rws in groups:
            body.append([gname, "", "", "", "", "", ""])
            for r in rws:
                body.append([r, "", "", "", "", "", ""])
        table_str = _grid_md([_ANALOG_H1, _ANALOG_H2], body, _ANALOG_MIN)
        idx = device_name.replace("A", "") if "A" in device_name else device_name
        crossref_id = f"analog{idx}"
        GENERATED_ANALOG_IDS.append(crossref_id)
        tables.append("\n\n" + table_str + "\n" + _caption_line(f"Проверка аналоговых каналов ({device_name})", crossref_id))
    return "\n".join(tables)

def _terminals_to_str(value):
    if isinstance(value, (list, tuple)):
        s = ", ".join(str(x) for x in value)
    else:
        s = "" if value is None else str(value)
    return s.replace("\u0425", "X").replace("\u0445", "x")


def render_discret_from_json(vars_data):
    cabinet_id = str(vars_data.get("id", ""))
    json_path = None
    if DISCRET_JSON_DIR.is_dir():
        for p in DISCRET_JSON_DIR.glob("*.json"):
            if cabinet_id in p.name:
                json_path = p
                break
    if not json_path or not json_path.is_file():
        print(f"[general_updater] WARNING: JSON дискретных {cabinet_id} не найден в {DISCRET_JSON_DIR}", file=sys.stderr)
        return ""
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            raw_data = json.load(f).get("data", {})
    except Exception as e:
        print(f"[general_updater] ERROR {json_path}: {e}", file=sys.stderr)
        return ""

    tables = []
    for device_name, device_data in raw_data.items():
        if not isinstance(device_data, dict) or not device_data:
            continue
        if isinstance(list(device_data.values())[0], dict):
            continue
        items = list(device_data.items())
        mid = (len(items) + 1) // 2
        col1, col2 = items[:mid], items[mid:]
        rows = []
        for i in range(max(len(col1), len(col2))):
            c1 = f"{col1[i][0]} \\\n{_terminals_to_str(col1[i][1])}" if i < len(col1) else ""
            c2 = f"{col2[i][0]} \\\n{_terminals_to_str(col2[i][1])}" if i < len(col2) else ""
            rows.append([c1, "", c2, ""])
        table_str = _grid_md([DISCRET_HEADERS], rows)
        crossref_id = f"discret_{device_name}"
        GENERATED_DISCRET_IDS.append(crossref_id)
        tables.append("\n\n" + table_str + "\n" + _caption_line(f"Пороги срабатывания дискретных входов ({device_name})", crossref_id))
    return "\n".join(tables)


def _expand_repeat(fname, vars_data, templates_dir):
    tables = []
    for idx, key in enumerate(REPEAT_BY_FILE[fname], start=1):
        code = vars_data.get(key)
        if not _nonempty(code):
            continue
        ctx = dict(vars_data)
        ctx["code_order"] = code
        ctx["code_hmi"] = vars_data.get(f"code_hmi{idx}") if idx > 1 else vars_data.get("code_hmi")
        ctx["label"] = f"A{idx}"
        ctx["N"] = str(idx)
        wb = _substitute_in_xlsx(templates_dir / fname, ctx)
        ws = _pick_data_sheet(wb)
        if "Info" in wb.sheetnames:
            for row in wb["Info"].iter_rows(min_row=1, max_col=2):
                k = str(row[0].value).strip().lower() if row[0].value else ""
                if k == "название таблицы":
                    if fname == "unit.xlsx":
                        row[1].value = f"Технические данные ЮНИТ (A{idx})"
                    elif fname == "hmi.xlsx":
                        row[1].value = f"Технические данные ИЧМ (A{idx}.1)"
                elif k in ("тэг", "tag", "id", "идентификатор"):
                    if fname == "unit.xlsx":
                        row[1].value = f"unit{idx}"
                    elif fname == "hmi.xlsx":
                        row[1].value = f"unit{idx}_1"
        try:
            md = _save_and_render(wb, ws)
        except Exception as e:
            print(f"[general_updater] ERROR {fname} N={idx}: {e}", file=sys.stderr)
            continue
        md, caps, found_ids = _split_trailing_captions(md)
        caption = caps[0] if caps else (_get_info_caption(wb) or "")
        fallback = f"unit{idx}" if fname == "unit.xlsx" else f"hmi{idx}"
        crossref_id = found_ids[0] if found_ids else fallback
        if caption:
            md += "\n" + _caption_line(caption, crossref_id)
        tables.append(md)
    return ("\n\n".join(tables) + "\n") if tables else ""


BLOCK_RE = re.compile(
    r"<!--\s*INSERT_TABLE:file=(?P<file>[^\s,>]+)(?:[^>]*)?\s*-->"
    r"(?P<body>.*?)"
    r"<!--\s*END_INSERT_TABLE\s*-->",
    re.DOTALL,
)


def _refs_phrase(ids, one, many):
    if not ids:
        return ""
    if len(ids) == 1:
        return f"{one} [@tbl:{ids[0]}]"
    return f"{many} [" + "; ".join(f"@tbl:{i}" for i in ids) + "]"


def _expand_refs(text, family, ids, one, many):
    # {{F_REFS}} — слово+список, {{F_WORD}} — слово, {{F_LIST}} — список;
    # если таблиц нет — строки с токенами удаляются целиком.
    if ids:
        text = text.replace("{{" + family + "_REFS}}", _refs_phrase(ids, one, many))
        text = text.replace("{{" + family + "_WORD}}", one if len(ids) == 1 else many)
        text = text.replace("{{" + family + "_LIST}}", "[" + "; ".join(f"@tbl:{i}" for i in ids) + "]")
    else:
        text = re.sub(r"^.*\{\{" + family + r"_(REFS|WORD|LIST)\}\}.*$", "", text, flags=re.MULTILINE)
    return text


def process_general_md(template_text, vars_data, templates_dir):
    def repl(m):
        fname = m.group("file")
        if fname in REPEAT_BY_FILE:
            return _expand_repeat(fname, vars_data, templates_dir)
        if fname == "discret.xlsx":
            return render_discret_from_json(vars_data)
        if fname == "analog.xlsx":
            return render_analog_from_json(vars_data)
        gate = GATE_BY_FILE.get(fname)
        if gate is not None and not _nonempty(vars_data.get(gate)):
            return ""
        xlsx_path = templates_dir / fname
        if not xlsx_path.is_file():
            return m.group(0)
        wb = _substitute_in_xlsx(xlsx_path, dict(vars_data))
        info_caption = _get_info_caption(wb)
        ws = _pick_data_sheet(wb)
        try:
            md = _save_and_render(wb, ws)
        except Exception as e:
            print(f"[general_updater] ERROR {fname}: {e}", file=sys.stderr)
            return m.group(0)
        md, caps, found_ids = _split_trailing_captions(md)
        caption = caps[0] if caps else info_caption
        crossref_id = found_ids[0] if found_ids else _slugify(caption or "table")
        if caption:
            md += "\n" + _caption_line(caption, crossref_id)
        return md + "\n"

    result = BLOCK_RE.sub(repl, template_text)
    result = _expand_refs(result, "DISCRET", GENERATED_DISCRET_IDS, "таблицу", "таблицы")
    result = _expand_refs(result, "ANALOG", GENERATED_ANALOG_IDS, "таблицу", "таблицы")
    result = re.sub(r"[ ]{2,}(?=\[)", " ", result)
    return result


def main():
    if len(sys.argv) != 5:
        print("Usage: general_updater.py <vars.json> <template_general.md> <out_general.md> <templates_dir>", file=sys.stderr)
        return 2
    vars_json, template_md, out_md, templates_dir = map(Path, sys.argv[1:])
    out_md.parent.mkdir(parents=True, exist_ok=True)
    vars_data = json.loads(vars_json.read_text(encoding="utf-8"))
    result = process_general_md(template_md.read_text(encoding="utf-8"), vars_data, templates_dir)
    out_md.write_text(result, encoding="utf-8", newline="\n")
    print(f"[general_updater] wrote {out_md}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
