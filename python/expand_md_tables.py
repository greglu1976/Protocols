#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Размножает таблицы в Markdown по числу заполненных слотов
и согласует число/падеж слова «устройство».

Формат пары «таблица + подпись» (подпись ПОД таблицей):

    +---+---+
    | … | … |
    +---+---+

    : Технические данные ЮНИТ <<REPEAT_UNIT>> {#tbl:tbl_unit}

После обработки:
    - суффикс (A1), (A2), … вставляется ПЕРЕД {#tbl:...};
    - ID делается уникальным: {#tbl:tbl_unit_1}, {#tbl:tbl_unit_2}, …

Нумерация таблиц («Таблица N.M – ») здесь НЕ делается — она
проставляется на этапе постобработки docx.

Маркеры:

  Таблицы:
    <<REPEAT_UNIT>>        — подпись таблицы ЮНИТ
    <<REPEAT_HMI>>         — подпись таблицы ИЧМ
    <<REPEAT_PRMPRD>>      — подпись таблицы ВЧ ПРМ/ПРД
                             (поддерживается и <<REPEAT_PRV>>)
    <<REPEAT_INSUL_UNIT>>  — подпись таблицы 7.x «Проверка сопротивления изоляции»
    <<CODE>>               — ячейка «Код заказа» (в таблицах ЮНИТ / ИЧМ / ПРВ)

  Согласование числа слова «устройство»:
    <<UNITS_GEN>>    — род. падеж:   «устройства»  / «устройств»
    <<UNITS_NOM>>    — им.  падеж:   «устройство»  / «устройства»
    <<UNITS_INSTR>>  — тв.  падеж:   «устройством» / «устройствами»
    <<UNITS_PREP>>   — пр.  падеж:   «устройстве»  / «устройствах»

Использование:
    python expand_md_tables.py <file.md> <vars.json>
"""
import json
import re
import sys
from pathlib import Path


# ─── Маркеры ────────────────────────────────────────────────────────────────

UNIT_MARKER        = "<<REPEAT_UNIT>>"
HMI_MARKER         = "<<REPEAT_HMI>>"
PRV_MARKERS        = ("<<REPEAT_PRV>>", "<<REPEAT_PRMPRD>>")  # оба варианта
INSUL_UNIT_MARKER  = "<<REPEAT_INSUL_UNIT>>"
CODE_MARKER        = "<<CODE>>"

UNITS_GEN_MARKER   = "<<UNITS_GEN>>"
UNITS_NOM_MARKER   = "<<UNITS_NOM>>"
UNITS_INSTR_MARKER = "<<UNITS_INSTR>>"
UNITS_PREP_MARKER  = "<<UNITS_PREP>>"

UNREPLACED_RE      = re.compile(r"<<[A-Z_]+>>")
TABLE_LINE_RE      = re.compile(r"^\s*[|+]")
CAPTION_LINE_RE    = re.compile(r"^:\s+")

# «{#tbl:xxx}» в самом конце строки (с опциональными пробелами перед ним)
CROSSREF_TAIL_RE   = re.compile(r'^(.*?)\s*\{#tbl:([^}]+)\}\s*$')


# ─── Слоты ──────────────────────────────────────────────────────────────────

def _slots(params, base):
    return [
        (params.get(base)       or "").strip(),
        (params.get(base + "2") or "").strip(),
        (params.get(base + "3") or "").strip(),
    ]


def _filled(params, base):
    return [(i + 1, v) for i, v in enumerate(_slots(params, base)) if v]


# ─── Согласование слова «устройство» ────────────────────────────────────────

def _ru_units(n, case="gen"):
    if case == "gen":
        return "устройства" if n == 1 else "устройств"
    if case == "nom":
        return "устройство" if n == 1 else "устройства"
    if case == "instr":
        return "устройством" if n == 1 else "устройствами"
    if case == "prep":
        return "устройстве" if n == 1 else "устройствах"
    raise ValueError(f"unknown case: {case}")


def _replace_marker_in_lines(lines, marker, text):
    for i, ln in enumerate(lines):
        if marker in ln:
            lines[i] = ln.replace(marker, text)


# ─── Парсинг и перерисовка таблицы ──────────────────────────────────────────

def _parse_table_cells(table_lines):
    rows = []
    for ln in table_lines:
        stripped = ln.strip()
        if stripped and set(stripped) <= set("+-=| \t"):
            continue
        if "|" in ln:
            parts = ln.split("|")
            if parts and parts[0].strip() == "":
                parts = parts[1:]
            if parts and parts[-1].strip() == "":
                parts = parts[:-1]
            rows.append([p.strip() for p in parts])
    return rows


def _render_grid_table(rows):
    if not rows:
        return []
    ncols = max(len(r) for r in rows)
    rows  = [list(r) + [""] * (ncols - len(r)) for r in rows]
    widths = [max(len(r[c]) for r in rows) for c in range(ncols)]
    widths = [max(w, 3) for w in widths]

    def hline(ch):
        return "+" + "+".join(ch * (w + 2) for w in widths) + "+"

    def row_line(cells):
        return "| " + " | ".join(c.ljust(w) for c, w in zip(cells, widths)) + " |"

    out = [hline("-"), row_line(rows[0]), hline("=")]
    for r in rows[1:]:
        out.append(row_line(r))
        out.append(hline("-"))
    return out


# ─── Поиск таблицы над подписью ─────────────────────────────────────────────

def _find_table_above(lines, cap_idx):
    j = cap_idx - 1
    if j >= 0 and lines[j].strip() == "":
        j -= 1
    end = j + 1
    while j >= 0 and TABLE_LINE_RE.match(lines[j]):
        j -= 1
    start = j + 1
    if start == end:
        return None
    return start, end


# ─── Правка подписи ─────────────────────────────────────────────────────────

def _caption_with_suffix(caption_line, suffix, suffix_id):
    """Вставляет суффикс ПЕРЕД {#tbl:...} и делает ID уникальным.
    Пример: ': Технические данные ЮНИТ {#tbl:tbl_unit}'
        → ': Технические данные ЮНИТ (A1) {#tbl:tbl_unit_1}'.
    """
    line = caption_line.rstrip()
    m = CROSSREF_TAIL_RE.match(line)
    if not m:
        return f"{line} {suffix}"
    head = m.group(1).rstrip()
    tid  = m.group(2)
    new_id = f"{tid}_{suffix_id}"
    return f"{head} {suffix} {{#tbl:{new_id}}}"


# ─── Размножение пары «таблица + подпись» ───────────────────────────────────

def expand_pair(lines, params, marker, base,
                suffix_fmt="(A{n})",
                fixed_suffix=None,
                fixed_suffix_id=None):
    """
    Размножает пару «таблица + подпись» по числу заполненных слотов.
    Маркер стоит в строке подписи (начинается с ':'), таблица — НАД ней.
    Суффикс добавляется ВСЕГДА: (A1), (A2), …, (A1.1), …
    Если слотов нет — пара удаляется целиком.
    """
    filled = _filled(params, base)
    n = len(lines)

    # ── Фаза 1: пары (tbl_start, tbl_end, cap_idx) ───────────────────────
    pairs = []
    i = 0
    while i < n:
        line = lines[i]
        if CAPTION_LINE_RE.match(line) and marker in line:
            tr = _find_table_above(lines, i)
            if tr is not None:
                pairs.append((tr[0], tr[1], i))
                i += 1
                continue
        i += 1

    if not pairs:
        return [ln.replace(marker, "") if marker in ln else ln for ln in lines]

    # ── Фаза 2: собираем результат ───────────────────────────────────────
    out = []
    cursor = 0
    for ts, te, cap_idx in pairs:
        chunk = lines[cursor:ts]
        while chunk and chunk[-1].strip() == "":
            chunk.pop()
        out.extend(chunk)

        if filled:
            table_lines = lines[ts:te]
            orig_rows   = _parse_table_cells(table_lines)
            cap_tpl     = lines[cap_idx]

            for num, code in filled:
                if out and out[-1].strip() != "":
                    out.append("")

                new_rows = [[c.replace(CODE_MARKER, code) for c in row]
                            for row in orig_rows]
                out.extend(_render_grid_table(new_rows))

                suffix    = fixed_suffix if fixed_suffix else suffix_fmt.format(n=num)
                suffix_id = str(fixed_suffix_id) if fixed_suffix_id else str(num)

                cap = cap_tpl.replace(marker, "").rstrip()
                cap = _caption_with_suffix(cap, suffix, suffix_id)
                out.append("")
                out.append(cap)

        cursor = cap_idx + 1

    tail = lines[cursor:]
    while tail and tail[0].strip() == "":
        tail.pop(0)
    if tail:
        if out and out[-1].strip() != "":
            out.append("")
        out.extend(tail)

    return [ln.replace(marker, "") if marker in ln else ln for ln in out]


# ─── Диагностика ────────────────────────────────────────────────────────────

def warn_unreplaced(lines):
    leftover = set()
    for ln in lines:
        for m in UNREPLACED_RE.finditer(ln):
            leftover.add(m.group(0))
    if leftover:
        print("[expand_md_tables] warning: необработанные маркеры:",
              ", ".join(sorted(leftover)), file=sys.stderr)


# ─── main ───────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) != 3:
        print("Usage: expand_md_tables.py <file.md> <vars.json>",
              file=sys.stderr)
        sys.exit(2)

    md_path, json_path = sys.argv[1], sys.argv[2]
    params = json.loads(Path(json_path).read_text(encoding="utf-8"))
    lines  = Path(md_path).read_text(encoding="utf-8").splitlines()

    n_unit = len(_filled(params, "code_order"))

    _replace_marker_in_lines(lines, UNITS_GEN_MARKER,   _ru_units(n_unit, "gen"))
    _replace_marker_in_lines(lines, UNITS_NOM_MARKER,   _ru_units(n_unit, "nom"))
    _replace_marker_in_lines(lines, UNITS_INSTR_MARKER, _ru_units(n_unit, "instr"))
    _replace_marker_in_lines(lines, UNITS_PREP_MARKER,  _ru_units(n_unit, "prep"))

    lines = expand_pair(lines, params, UNIT_MARKER, "code_order",  suffix_fmt="(A{n})")
    lines = expand_pair(lines, params, HMI_MARKER,  "code_hmi",    suffix_fmt="(A{n}.1)")
    for prv in PRV_MARKERS:
        lines = expand_pair(lines, params, prv, "code_prmprd",
                            fixed_suffix="(A2)", fixed_suffix_id=1)
    lines = expand_pair(lines, params, INSUL_UNIT_MARKER, "code_order",
                        suffix_fmt="(A{n})")

    warn_unreplaced(lines)

    Path(md_path).write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()