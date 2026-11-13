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

GENERATED_DISCRET_IDS = []
GENERATED_ANALOG_IDS = []
GENERATED_INSUL_IDS = []
GENERATED_OUTPUTS_IDS = []  # <--- Новая переменная

DISCRET_HEADERS = (
    "Наименование задействованного входа", "U ср / U в, В",
    "Наименование задействованного входа", "U ср / U в, В",
)

OUTPUTS_HEADERS = (
    "Наименование выходного реле", "Состояние",
    "Наименование выходного реле", "Состояние",
)

FILLER = "\u2800"  # пустой символ Брайля: не пробельный для конвертера,
                   # но невидим в Word — мешает объединению пустых ячеек

INSUL_HEADERS = ("№", "Наименование цепи", "Объединяемые клеммы шкафа")

DEFAULT_ANALOG_GROUPS = ["№", "Измерительный канал", "Входное значение", "Замер", "Погрешность, %"]
DEFAULT_ANALOG_SUB = ["", "", "", "I(U), А(В)", "∠φ, град.", "I(U), А(В)", "∠φ, град.", "Амплитудная", "Угловая"]


def _nonempty(v):
    if v is None:
        return False
    if isinstance(v, str):
        return bool(v.strip())
    if isinstance(v, bool):
        return v
    return True


def _sanitize_id(raw):
    s = re.sub(r"[^0-9A-Za-zА-Яа-яЁё_]+", "_", str(raw))
    s = s.strip("_") or "table"
    return s


def _caption_line(caption, crossref_id):
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


def _find_info_sheet_name(wb):
    for name in wb.sheetnames:
        if name.lower() == "info":
            return name
    return None


def _get_info_caption(wb):
    info_name = _find_info_sheet_name(wb)
    if info_name is None:
        return None
    for row in wb[info_name].iter_rows(min_row=1, max_col=2):
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
                if v.endswith(" \\"):
                    clean_v = v[:-2]
                    spaces = max(0, widths[i] - len(clean_v) - 2)
                    parts.append(" " + clean_v + " " * spaces + " \\ ")
                else:
                    spaces = max(0, widths[i] - len(v))
                    parts.append(" " + v + " " * spaces + " ")
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


def _read_analog_headers(path):
    if not path or not path.is_file():
        return DEFAULT_ANALOG_GROUPS, DEFAULT_ANALOG_SUB
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb['Лист1'] if 'Лист1' in wb.sheetnames else wb.worksheets[0]
        def get_val(r, c):
            for rng in ws.merged_cells.ranges:
                if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
                    v = ws.cell(rng.min_row, rng.min_col).value
                    return str(v).strip() if v is not None else ""
            v = ws.cell(r, c).value
            return str(v).strip() if v is not None else ""

        groups = [get_val(1, 1), get_val(1, 2), get_val(1, 4), get_val(1, 6), get_val(1, 8)]
        sub = [get_val(2, c) for c in range(1, 10)]
        sub[0] = sub[1] = sub[2] = ""
        for i in range(5):
            if not groups[i]:
                groups[i] = DEFAULT_ANALOG_GROUPS[i]
        for i in range(9):
            if not sub[i] and DEFAULT_ANALOG_SUB[i]:
                sub[i] = DEFAULT_ANALOG_SUB[i]
        return groups, sub
    except Exception as e:
        print(f"[general_updater] WARN: Не удалось прочитать шапку из {path}: {e}", file=sys.stderr)
        return DEFAULT_ANALOG_GROUPS, DEFAULT_ANALOG_SUB


def _make_analog_grid_table(rows, groups, sub):
    w0, w1, w2 = 4, 7, 30
    for r in rows:
        for ln in r[1].split("\n"):
            w2 = max(w2, len(ln))
    w3 = max(17, len(sub[3]))
    w4 = max(9,  len(sub[4]))
    w5 = max(17, len(sub[5]))
    w6 = max(9,  len(sub[6]))
    w7 = max(20, len(sub[7]))
    w8 = max(9,  len(sub[8]))
    pad = [w + 2 for w in (w0, w1, w2, w3, w4, w5, w6, w7, w8)]

    def full_border(ch="-"):
        return "+" + "+".join(ch * p for p in pad) + "+"

    def group_row(texts):
        spans = [pad[0], pad[1] + 1 + pad[2], pad[3] + 1 + pad[4], pad[5] + 1 + pad[6], pad[7] + 1 + pad[8]]
        parts = []
        for w, t in zip(spans, texts):
            t = str(t).strip()
            parts.append(t + " " * max(0, w - len(t)))
        return "|" + "|".join(parts) + "|"

    def cell_row(cells9, fill_empty=False):
        parts = []
        for p, t in zip(pad, cells9):
            t = str(t)
            if fill_empty and not t:
                t = FILLER
            parts.append(t + " " * max(0, p - len(t)))
        return "|" + "|".join(parts) + "|"

    md = []
    md.append(full_border("-"))
    md.append(group_row(groups))
    md.append(full_border("-"))
    md.append(cell_row(sub))
    md.append(full_border("="))

    for i, r in enumerate(rows, start=1):
        for j, ln in enumerate(r[1].split("\n")):
            md.append(cell_row(
                [str(i) if j == 0 else "",
                 r[0] if j == 0 else "",
                 ln] + [""] * 6,
                fill_empty=True,
            ))
        md.append(full_border("-"))

    return "\n".join(md)


def is_analog_group(name):
    n = name.lower()
    return "переменного" in n or "напряжения" in n


def extract_channels(data, path=None):
    if path is None:
        path = []
    channels = []
    for k, v in data.items():
        if isinstance(v, list):
            channels.append((path, k, v))
        elif isinstance(v, dict):
            channels.extend(extract_channels(v, path + [k]))
    return channels


def render_analog_from_json(json_data, analog_path):
    groups, sub = _read_analog_headers(analog_path)
    tables = []
    for device_name, device_data in json_data.get("data", {}).items():
        if device_name == "Группы цепей шкафа":
            continue
        analog_rows = []
        for group_name, group_data in device_data.items():
            if not is_analog_group(group_name):
                continue
            channels = extract_channels(group_data)
            for path, ch_name, terminals in channels:
                parts = list(path) if path else [group_name]
                term_str = ", ".join(str(t) for t in terminals)
                parts.append(term_str)
                cell_text = " \\\n".join(parts)
                base = ch_name.strip("«»") if ch_name else ""
                designator = f"«{base}»" if base else ""
                analog_rows.append([designator, cell_text])
        if analog_rows:
            table_str = _make_analog_grid_table(analog_rows, groups, sub)
            crossref_id = _sanitize_id(f"analog_{device_name}")
            GENERATED_ANALOG_IDS.append(crossref_id)
            caption = f"\n\n: Проверка аналоговых каналов ({device_name}) {{#tbl:{crossref_id}}}\n\n"
            tables.append(f"\n\n{table_str}{caption}\n\n")
    return "\n".join(tables)


def render_discret_from_json(json_data):
    tables = []
    for device_name, device_data in json_data.get("data", {}).items():
        if device_name == "Группы цепей шкафа":
            continue

        def find_dv_leaves(data):
            found = []
            if isinstance(data, dict):
                for k, v in data.items():
                    if isinstance(v, list) and k.startswith("ДВ"):
                        found.append((k, v))
                    elif isinstance(v, dict):
                        found.extend(find_dv_leaves(v))
            return found

        discret_items = []
        for group_name, group_data in device_data.items():
            if not group_name.startswith("Цепи оперативного тока"):
                continue
            for ch_name, terminals in find_dv_leaves(group_data):
                term_str = ", ".join(str(t) for t in terminals)
                discret_items.append(f"{ch_name} \\\n{term_str}")

        if not discret_items:
            continue

        mid = (len(discret_items) + 1) // 2
        col1, col2 = discret_items[:mid], discret_items[mid:]
        rows = []
        for i in range(max(len(col1), len(col2))):
            c1 = col1[i] if i < len(col1) else ""
            c2 = col2[i] if i < len(col2) else ""
            rows.append([c1, "", c2, ""])

        table_str = _grid_md([DISCRET_HEADERS], rows)
        crossref_id = _sanitize_id(f"discret_{device_name}")
        GENERATED_DISCRET_IDS.append(crossref_id)
        caption = f"\n\n: Пороги срабатывания дискретных входов ({device_name}) {{#tbl:{crossref_id}}}\n\n"
        tables.append(f"\n\n{table_str}{caption}\n\n")
    return "\n".join(tables)


def render_outputs_from_json(json_data):
    tables = []
    for device_name, device_data in json_data.get("data", {}).items():
        if device_name == "Группы цепей шкафа":
            continue

        output_items = []
        for group_name, group_data in device_data.items():
            if not group_name.startswith("Выходные цепи"):
                continue
            if isinstance(group_data, dict):
                for ch_name, terminals in group_data.items():
                    term_str = ", ".join(str(t) for t in terminals)
                    output_items.append(f"{ch_name} \\\n{term_str}")

        if not output_items:
            continue

        mid = (len(output_items) + 1) // 2
        col1, col2 = output_items[:mid], output_items[mid:]
        rows = []
        for i in range(max(len(col1), len(col2))):
            c1 = col1[i] if i < len(col1) else ""
            c2 = col2[i] if i < len(col2) else ""
            rows.append([c1, "", c2, ""])

        table_str = _grid_md([OUTPUTS_HEADERS], rows)
        crossref_id = _sanitize_id(f"outputs_{device_name}")
        GENERATED_OUTPUTS_IDS.append(crossref_id)
        caption = f"\n\n: Срабатывания дискретных выходов ({device_name}) {{#tbl:{crossref_id}}}\n\n"
        tables.append(f"\n\n{table_str}{caption}\n\n")
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

        info_name = _find_info_sheet_name(wb)
        if info_name is not None:
            ws_info = wb[info_name]
            for row in ws_info.iter_rows(min_row=1, max_col=2):
                k = str(row[0].value).strip().lower() if row[0].value else ""
                if k == 'название таблицы':
                    if fname == "unit.xlsx":
                        row[1].value = f"Технические данные ЮНИТ (A{idx})"
                    elif fname == "hmi.xlsx":
                        row[1].value = f"Технические данные ИЧМ (A{idx}.1)"
                elif k in ('тэг', 'tag', 'id', 'идентификатор'):
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

    if not tables:
        return ""
    return "\n\n".join(tables) + "\n"


def render_insul_from_json(json_data):
    groups = json_data.get("data", {}).get("Группы цепей шкафа")

    if isinstance(groups, dict):
        items = [(str(k), str(v)) for k, v in groups.items()]
    elif isinstance(groups, list):
        items = [(str(x), "") for x in groups]
    else:
        items = []

    if not items:
        print("[general_updater] WARNING: 'Группы цепей шкафа' не найдены в JSON", file=sys.stderr)
        return ""

    rows = []
    for i, (name, terminals) in enumerate(items, start=1):
        rows.append([str(i), name, terminals])

    table_str = _grid_md([list(INSUL_HEADERS)], rows)
    GENERATED_INSUL_IDS.append("groups")
    caption = "\n\n: Объединенные группы цепей шкафа {#tbl:groups}\n\n"
    return f"\n\n{table_str}{caption}\n\n"


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
    if ids:
        text = text.replace("{{" + family + "_REFS}}", _refs_phrase(ids, one, many))
        text = text.replace("{{" + family + "_WORD}}", one if len(ids) == 1 else many)
        text = text.replace("{{" + family + "_LIST}}", "[" + "; ".join(f"@tbl:{i}" for i in ids) + "]")
    else:
        text = text.replace("{{" + family + "_REFS}}", "")
        text = text.replace("{{" + family + "_WORD}}", "")
        text = text.replace("{{" + family + "_LIST}}", "")
        text = re.sub(r"\s+\.", ".", text)
        text = re.sub(r"\s+,", ",", text)
        text = re.sub(r"\(\s*\)", "", text)
    return text


def _find_analog_json(cabinet_dirs, cabinet_id):
    for d in cabinet_dirs:
        if not d.is_dir():
            continue
        exact = d / f"{cabinet_id}.json"
        if exact.is_file():
            return exact
        for p in d.glob("*.json"):
            if p.name == "sections.json":
                continue
            if cabinet_id in p.name:
                return p
        candidates = [p for p in d.glob("*.json") if p.name != "sections.json"]
        if len(candidates) == 1:
            return candidates[0]
    return None


def process_general_md(template_text, vars_data, templates_dir, cabinet_dirs):
    GENERATED_DISCRET_IDS.clear()
    GENERATED_ANALOG_IDS.clear()
    GENERATED_INSUL_IDS.clear()
    GENERATED_OUTPUTS_IDS.clear()

    cabinet_id = str(vars_data.get("id", ""))
    json_path = _find_analog_json(cabinet_dirs, cabinet_id)

    if not json_path:
        print(f"[general_updater] WARNING: JSON для '{cabinet_id}' не найден в папках: {[str(d) for d in cabinet_dirs]}", file=sys.stderr)
        json_data = {}
    else:
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                json_data = json.load(f)
            print(f"[general_updater] INFO: Успешно прочитан JSON: {json_path}")
        except Exception as e:
            print(f"[general_updater] ERROR reading {json_path}: {e}", file=sys.stderr)
            json_data = {}

    # ── Настройка связи (Зависит от последнего code_order) ──
    last_code = ""
    for key in ("code_order", "code_order2", "code_order3"):
        val = vars_data.get(key)
        if val and str(val).strip():
            last_code = str(val).strip()

    comms_info = ""
    if "С01" in last_code or "C01" in last_code:
        comms_info = "Устройство РЗА содержит порты связи (Ethernet – 4 шт.), располагающиеся на модуле центрального процессора."
    elif "С04" in last_code or "C04" in last_code:
        comms_info = "Устройство РЗА содержит порты связи (Ethernet – 4 шт.; C37.94 - 2шт.; RS485 - 2шт.), располагающиеся на модуле центрального процессора."
    else:
        comms_info = "Устройство РЗА содержит порты связи (Ethernet – 4 шт.), располагающиеся на модуле центрального процессора."

    template_text = template_text.replace("{{COMMS_PORT_INFO}}", comms_info)

    def repl(m):
        fname = m.group("file")
        if fname in REPEAT_BY_FILE:
            return _expand_repeat(fname, vars_data, templates_dir)
        if fname == "discret.xlsx":
            return render_discret_from_json(json_data)
        if fname == "analog.xlsx":
            return render_analog_from_json(json_data, templates_dir / fname)
        if fname == "groups.xlsx":
            return render_insul_from_json(json_data)
        if fname == "outputs.xlsx":
            return render_outputs_from_json(json_data)
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
    result = _expand_refs(result, "OUTPUTS", GENERATED_OUTPUTS_IDS, "таблицу", "таблицы")
    result = re.sub(r"[ ]{2,}(?=\[)", " ", result)
    return result


def main():
    if len(sys.argv) != 5:
        print("Usage: general_updater.py <vars.json> <template_general.md> <out_general.md> <templates_dir>", file=sys.stderr)
        return 2

    vars_json, template_md, out_md, templates_dir = map(Path, sys.argv[1:])
    out_md.parent.mkdir(parents=True, exist_ok=True)
    vars_data = json.loads(vars_json.read_text(encoding="utf-8"))

    cabinet_dirs = [
        templates_dir.parent / "cabinet",
        out_md.parent,
        Path("cabinet"),
        Path(".")
    ]

    result = process_general_md(
        template_md.read_text(encoding="utf-8"),
        vars_data,
        templates_dir,
        cabinet_dirs,
    )
    out_md.write_text(result, encoding="utf-8", newline="\n")

    print(f"[general_updater] wrote {out_md}")
    return 0


if __name__ == "__main__":
    sys.exit(main())