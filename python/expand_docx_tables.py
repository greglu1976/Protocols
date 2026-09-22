# -*- coding: utf-8 -*-
"""Размножает таблицы «Технические данные ЮНИТ / ИЧМ / ВЧ ПРМ/ПРД» в docx
по числу заполненных слотов.

В шаблоне должны быть маркеры:
  <<REPEAT_UNIT>>  — в абзаце-подписи таблицы ЮНИТ
  <<REPEAT_HMI>>   — в абзаце-подписи таблицы ИЧМ
  <<REPEAT_PRV>>   — в абзаце-подписи таблицы ВЧ ПРМ/ПРД
  <<CODE>>         — в ячейке «Код заказа» соответствующей таблицы

Правила:
  - ЮНИТ: по таблице на каждый заполненный слот code_order[2,3],
          суффиксы (A1), (A2), (A3) — всегда, даже если слот один;
          на подписях со второй по счёту — интервал «Перед: 6 пт»;
  - ИЧМ:  аналогично, но суффиксы (A1.1), (A2.1), (A3.1) — всегда;
  - ВЧ ПРМ/ПРД: один слот code_prmprd. Если пусто — подпись и таблица
          удаляются целиком; если заполнено — одна таблица с фиксированным
          суффиксом (A2);
  - после вставки скрипт перенумеровывает «Таблица X.Y» по порядку
    (2.1, 2.2, 2.3, …).

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


UNIT_MARKER = "<<REPEAT_UNIT>>"
HMI_MARKER  = "<<REPEAT_HMI>>"
PRV_MARKER  = "<<REPEAT_PRV>>"
CODE_MARKER = "<<CODE>>"

TABLE_NUM_RE = re.compile(r"(Таблица\s+)(\d+)\.(\d+)")

# Интервал «Перед» для подписей (A2), (A3), … — 6 пт.
SUFFIX_SPACE_BEFORE_PT = 6


# ─── Обход блоков в порядке следования ────────────────────────────────────────

def iter_blocks(doc):
    """Yield ('p', Paragraph) или ('t', Table) в порядке появления в body."""
    body = doc.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield 'p', Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield 't', Table(child, doc)


# ─── Слоты кодов ──────────────────────────────────────────────────────────────

def _slots(params, base):
    return [
        (params.get(base)       or "").strip(),
        (params.get(base + "2") or "").strip(),
        (params.get(base + "3") or "").strip(),
    ]


def _filled(params, base):
    """[(номер_слота_1..3, код), …] — только заполненные."""
    return [(i + 1, v) for i, v in enumerate(_slots(params, base)) if v]


# ─── Работа с текстом ─────────────────────────────────────────────────────────

def _replace_in_paragraph(paragraph, old, new):
    """Заменяет old→new по всем runs абзаца, сохраняя форматирование
    первого run."""
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


# ─── Размножение пары «подпись + таблица» ────────────────────────────────────

def expand_pair(doc, params, marker, base,
                suffix_fmt="(A{n})",
                fixed_suffix=None):
    """
    suffix_fmt   — Python-format-строка для суффикса в подписи.
                    UNIT: "(A{n})"    → (A1), (A2), (A3)
                    HMI:  "(A{n}.1)"  → (A1.1), (A2.1), (A3.1)
    fixed_suffix — если задан, используется вместо suffix_fmt (например,
                    "(A2)" для ВЧ ПРМ/ПРД). Суффикс тогда одинаков для
                    всех таблиц этой группы.

    Суффикс добавляется ВСЕГДА, даже если слот один.

    Если в документе найден абзац с маркером, ищется следующая за ним
    таблица, и пара «подпись + таблица» клонируется столько раз, сколько
    заполнено слотов в params[base + ""|"2"|"3"].
    """
    blocks = list(iter_blocks(doc))
    i = 0
    while i < len(blocks):
        kind, obj = blocks[i]
        if kind != 'p' or marker not in obj.text:
            i += 1
            continue

        # Следующая таблица — цель
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
            # Нет значений — удаляем и подпись, и таблицу
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

            # Суффикс: fixed_suffix имеет приоритет; иначе — по номеру слота
            if fixed_suffix:
                suffix = fixed_suffix
            else:
                suffix = suffix_fmt.format(n=n)
            _append_suffix(new_cap, suffix)

            # Интервал «Перед» у всех подписей, кроме первой
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


# ─── main ────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) != 3:
        print("Usage: expand_docx_tables.py <docx> <vars.json>", file=sys.stderr)
        sys.exit(2)

    docx_path, json_path = sys.argv[1], sys.argv[2]
    with open(json_path, encoding="utf-8") as f:
        params = json.load(f)

    doc = Document(docx_path)

    expand_pair(doc, params, UNIT_MARKER, "code_order",  suffix_fmt="(A{n})")
    expand_pair(doc, params, HMI_MARKER,  "code_hmi",    suffix_fmt="(A{n}.1)")
    expand_pair(doc, params, PRV_MARKER,  "code_prmprd", fixed_suffix="(A2)")

    renumber_tables(doc)

    doc.save(docx_path)


if __name__ == "__main__":
    main()