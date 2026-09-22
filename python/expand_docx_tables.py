# -*- coding: utf-8 -*-
"""Размножает таблицы в docx по числу заполненных слотов и согласует
число/падеж слова «устройство».

Маркеры в шаблоне (title_page.docx):

  Таблицы:
    <<REPEAT_UNIT>>        — подпись таблицы ЮНИТ
    <<REPEAT_HMI>>         — подпись таблицы ИЧМ
    <<REPEAT_PRV>>         — подпись таблицы ВЧ ПРМ/ПРД
                             (поддерживается и написание <<REPEAT_PRIV>>)
    <<REPEAT_INSUL_UNIT>>  — подпись таблицы 7.x «Проверка сопротивления изоляции»
    <<CODE>>               — ячейка «Код заказа» (в таблицах ЮНИТ / ИЧМ / ПРВ)

  Ссылки:
    <<REFS_INSUL_UNIT>>    — «в таблицу 7.1» / «в таблицы 7.1 и 7.2» / …
    <<REFS_INSUL_CAB>>     — «в таблицу 7.N+1» (для одиночной таблицы 7.2)

  Согласование числа слова «устройство» (по количеству слотов code_order):
    <<UNITS_GEN>>    — род. падеж:   «устройства»  / «устройств»
    <<UNITS_NOM>>    — им.  падеж:   «устройство»  / «устройства»
    <<UNITS_INSTR>>  — тв.  падеж:   «устройством» / «устройствами»
    <<UNITS_PREP>>   — пр.  падеж:   «устройстве»  / «устройствах»

Правила:
  - ЮНИТ: N таблиц по числу заполненных code_order[2,3]; суффиксы (A1),(A2),(A3)
          добавляются всегда; на подписях со второй — «Перед: 6 пт»;
  - ИЧМ:  аналогично, суффиксы (A1.1),(A2.1),(A3.1);
  - ВЧ ПРМ/ПРД: один слот code_prmprd; если пусто — таблица и подпись
          удаляются; если заполнено — одна таблица с суффиксом (A2);
  - Раздел 7.1: N таблиц по числу заполненных code_order[2,3];
          маркер <<REFS_INSUL_UNIT>> заменяется на «в таблицу 7.1» /
          «в таблицы 7.1 и 7.2» / «в таблицы 7.1, 7.2 и 7.3»;
  - Раздел 7.2: одна таблица, номер = N+1; маркер <<REFS_INSUL_CAB>>
          заменяется на «в таблицу 7.N+1»;
  - Перенумерация «Таблица X.Y» по порядку появления внутри каждой секции X.

Использование:
    python expand_docx_tables.py <docx> <vars.json>
"""
import json
import re
import sys
from copy import deepcopy

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import Pt
from docx.table import Table
from docx.text.paragraph import Paragraph


UNIT_MARKER        = "<<REPEAT_UNIT>>"
HMI_MARKER         = "<<REPEAT_HMI>>"
PRV_MARKERS        = ("<<REPEAT_PRV>>", "<<REPEAT_PRIV>>")  # оба варианта
INSUL_UNIT_MARKER  = "<<REPEAT_INSUL_UNIT>>"
REFS_INSUL_MARKER  = "<<REFS_INSUL_UNIT>>"
REFS_INSUL_CAB_MARKER = "<<REFS_INSUL_CAB>>"
CODE_MARKER        = "<<CODE>>"

UNITS_GEN_MARKER   = "<<UNITS_GEN>>"
UNITS_NOM_MARKER   = "<<UNITS_NOM>>"
UNITS_INSTR_MARKER = "<<UNITS_INSTR>>"
UNITS_PREP_MARKER  = "<<UNITS_PREP>>"

TABLE_NUM_RE     = re.compile(r"(Таблица\s+)(\d+)\.(\d+)")
UNREPLACED_RE    = re.compile(r"<<[A-Z_]+>>")

SUFFIX_SPACE_BEFORE_PT = 6


# ─── Обход блоков ────────────────────────────────────────────────────────────

def iter_blocks(doc):
    """Yield ('p', Paragraph) или ('t', Table) в порядке появления в body."""
    body = doc.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield 'p', Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield 't', Table(child, doc)


# ─── Слоты ───────────────────────────────────────────────────────────────────

def _slots(params, base):
    return [
        (params.get(base)       or "").strip(),
        (params.get(base + "2") or "").strip(),
        (params.get(base + "3") or "").strip(),
    ]


def _filled(params, base):
    """[(номер_слота_1..3, код), …] — только заполненные."""
    return [(i + 1, v) for i, v in enumerate(_slots(params, base)) if v]


# ─── Работа с текстом ────────────────────────────────────────────────────────

def _replace_in_paragraph(paragraph, old, new):
    """Заменяет old→new, собирая текст абзаца в первый run."""
    if old in paragraph.text:
        text = paragraph.text.replace(old, new)
        if paragraph.runs:
            paragraph.runs[0].text = text
            for r in paragraph.runs[1:]:
                r.text = ""
        else:
            paragraph.add_run(text)


def _clear_marker(paragraph, marker):
    _replace_in_paragraph(paragraph, marker, "")


def _set_code_in_table(table, code):
    """Заменяет <<CODE>> на code во всех ячейках таблицы."""
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                if CODE_MARKER in p.text:
                    _replace_in_paragraph(p, CODE_MARKER, code)


def _append_suffix(paragraph, suffix):
    """Дописывает ' (A1)' в конец абзаца без двойного пробела."""
    for r in reversed(paragraph.runs):
        if r.text.strip() == "":
            r.text = ""
            continue
        r.text = r.text.rstrip() + f" {suffix}"
        return
    paragraph.add_run(f" {suffix}")


# ─── Ссылки «в таблицу 7.1» / «в таблицы 7.1 и 7.2» / … ─────────────────────

def _ru_refs(nums, section=7):
    """
    nums = [1]      → 'в таблицу 7.1'
    nums = [1,2]    → 'в таблицы 7.1 и 7.2'
    nums = [1,2,3]  → 'в таблицы 7.1, 7.2 и 7.3'
    nums = []       → ''
    """
    if not nums:
        return ""
    prefixed = [f"{section}.{n}" for n in nums]
    if len(prefixed) == 1:
        return f"в таблицу {prefixed[0]}"
    if len(prefixed) == 2:
        return f"в таблицы {prefixed[0]} и {prefixed[1]}"
    return "в таблицы " + ", ".join(prefixed[:-1]) + " и " + prefixed[-1]


def replace_marker(doc, marker, text):
    """Заменяет marker на text во всех абзацах документа (body)."""
    for p in doc.paragraphs:
        if marker in p.text:
            _replace_in_paragraph(p, marker, text)


# ─── Согласование слова «устройство» с числом слотов ────────────────────────

def _ru_units(n, case="gen"):
    """
    n    — количество устройств (1, 2, 3, …)
    case — падеж: 'gen' (род.), 'nom' (им.), 'instr' (тв.), 'prep' (пр.)

    Возвращает:
      gen   → «устройства»  / «устройств»
      nom   → «устройство»  / «устройства»
      instr → «устройством» / «устройствами»
      prep  → «устройстве»  / «устройствах»
    """
    if case == "gen":
        return "устройства" if n == 1 else "устройств"
    if case == "nom":
        return "устройство" if n == 1 else "устройства"
    if case == "instr":
        return "устройством" if n == 1 else "устройствами"
    if case == "prep":
        return "устройстве" if n == 1 else "устройствах"
    raise ValueError(f"unknown case: {case}")


# ─── Размножение пары «подпись + таблица» ────────────────────────────────────

def expand_pair(doc, params, marker, base,
                suffix_fmt="(A{n})",
                fixed_suffix=None):
    """
    suffix_fmt   — Python-format-строка для суффикса в подписи:
                     "(A{n})"   → (A1), (A2), (A3)
                     "(A{n}.1)" → (A1.1), (A2.1), (A3.1)
    fixed_suffix — если задан, используется вместо suffix_fmt (например,
                   "(A2)" для ВЧ ПРМ/ПРД).

    Суффикс добавляется ВСЕГДА, даже если слот один.
    """
    blocks = list(iter_blocks(doc))
    i = 0
    while i < len(blocks):
        kind, obj = blocks[i]
        if kind != 'p' or marker not in obj.text:
            i += 1
            continue

        table = None
        for j in range(i + 1, len(blocks)):
            if blocks[j][0] == 't':
                table = blocks[j][1]
                break

        filled = _filled(params, base)

        if table is None:
            _clear_marker(obj, marker)
            blocks = list(iter_blocks(doc))
            i += 1
            continue

        if not filled:
            obj._element.getparent().remove(obj._element)
            table._element.getparent().remove(table._element)
            blocks = list(iter_blocks(doc))
            i = 0
            continue

        cap_el = obj._element
        tbl_el = table._element

        for idx, (n, code) in enumerate(filled):
            new_cap_el = deepcopy(cap_el)
            new_tbl_el = deepcopy(tbl_el)

            cap_el.addprevious(new_cap_el)
            cap_el.addprevious(new_tbl_el)

            new_cap = Paragraph(new_cap_el, obj._parent)
            new_tbl = Table(new_tbl_el, table._parent)

            _clear_marker(new_cap, marker)
            suffix = fixed_suffix if fixed_suffix else suffix_fmt.format(n=n)
            _append_suffix(new_cap, suffix)

            if idx > 0:
                new_cap.paragraph_format.space_before = Pt(SUFFIX_SPACE_BEFORE_PT)

            _set_code_in_table(new_tbl, code)

        cap_el.getparent().remove(cap_el)
        tbl_el.getparent().remove(tbl_el)

        blocks = list(iter_blocks(doc))
        i = 0


# ─── Перенумерация «Таблица X.Y» ─────────────────────────────────────────────

def renumber_tables(doc):
    """Перенумеровывает все 'Таблица X.Y' по порядку появления,
    счётчик сбрасывается на каждой секции X."""
    counters = {}
    for p in doc.paragraphs:
        m = TABLE_NUM_RE.search(p.text)
        if not m:
            continue
        section = int(m.group(2))
        counters[section] = counters.get(section, 0) + 1
        new_num = f"Таблица {section}.{counters[section]}"
        _replace_in_paragraph(p, m.group(0), new_num)


# ─── Диагностика необработанных маркеров ────────────────────────────────────

def warn_unreplaced(doc):
    leftover = set()
    for p in doc.paragraphs:
        for m in UNREPLACED_RE.finditer(p.text):
            leftover.add(m.group(0))
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for m in UNREPLACED_RE.finditer(p.text):
                        leftover.add(m.group(0))
    if leftover:
        print("[expand_docx_tables] warning: необработанные маркеры:",
              ", ".join(sorted(leftover)), file=sys.stderr)


# ─── main ────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) != 3:
        print("Usage: expand_docx_tables.py <docx> <vars.json>", file=sys.stderr)
        sys.exit(2)

    docx_path, json_path = sys.argv[1], sys.argv[2]
    with open(json_path, encoding="utf-8") as f:
        params = json.load(f)

    doc = Document(docx_path)

    # Количество устройств = количество заполненных code_order-слотов
    n_unit = len(_filled(params, "code_order"))

    # ── Согласование «устройство»/«устройства»/«устройств» ────────────────
    replace_marker(doc, UNITS_GEN_MARKER,   _ru_units(n_unit, "gen"))
    replace_marker(doc, UNITS_NOM_MARKER,   _ru_units(n_unit, "nom"))
    replace_marker(doc, UNITS_INSTR_MARKER, _ru_units(n_unit, "instr"))
    replace_marker(doc, UNITS_PREP_MARKER,  _ru_units(n_unit, "prep"))

    # ── Раздел 2: ЮНИТ / ИЧМ / ВЧ ПРМ/ПРД ─────────────────────────────────
    expand_pair(doc, params, UNIT_MARKER, "code_order",  suffix_fmt="(A{n})")
    expand_pair(doc, params, HMI_MARKER,  "code_hmi",    suffix_fmt="(A{n}.1)")
    for prv in PRV_MARKERS:
        expand_pair(doc, params, prv, "code_prmprd", fixed_suffix="(A2)")

    # ── Раздел 7.1: размножение таблиц изоляции по code_order ─────────────
    replace_marker(doc, REFS_INSUL_MARKER,
                   _ru_refs(list(range(1, n_unit + 1)), section=7))
    if n_unit > 0:
        replace_marker(doc, REFS_INSUL_CAB_MARKER,
                       f"в таблицу 7.{n_unit + 1}")

    expand_pair(doc, params, INSUL_UNIT_MARKER, "code_order", suffix_fmt="(A{n})")

    # ── Финальная перенумерация и диагностика ─────────────────────────────
    renumber_tables(doc)
    warn_unreplaced(doc)
    doc.save(docx_path)


if __name__ == "__main__":
    main()