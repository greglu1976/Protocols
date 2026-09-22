# -*- coding: utf-8 -*-
"""Читает строку из Excel по id и выгружает переменные кабинета.

Пишет два файла:
  <out.mk>    — NAME := … , CODE := … и т.д. для -include в Makefile;
  <out.json>  — те же значения + производные переменные для md_subst/inject.

Использование:
    python export_vars.py <id> <excel> <out.mk>
"""
import json
import sys
from datetime import datetime

import openpyxl


# ─── Утилиты для значений ─────────────────────────────────────────────────────

def fmt_date(value):
    """datetime или 'YYYY-MM-DD…' → 'DD.MM.YYYY'. Остальное — как есть."""
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    s = str(value).strip()
    if len(s) >= 10 and s[4] == "-" and s[7] == "-":
        y, m, d = s[:10].split("-")
        return f"{d}.{m}.{y}"
    return s


def normalize_description(text):
    """Литерал \\n (обратный слэш + n) превращаем в настоящий перевод строки."""
    s = str(text)
    s = s.replace("\\n", "\n")
    s = s.replace("\r\n", "\n")
    s = s.replace("\r", "\n")
    s = "\n".join(part.strip() for part in s.split("\n"))
    return s


def make_escape(s):
    """Экранирование для .mk-значения."""
    return s.replace("\\", "\\\\").replace("$", "$$").replace("#", "\\#")


def _to_mk_value(value):
    """Схлопывает все переводы/табы в пробелы — .mk-значение однострочное."""
    return (value
            .replace("\r\n", " ").replace("\r", " ")
            .replace("\n", " ").replace("\t", " "))


# ─── Слоты кодов ──────────────────────────────────────────────────────────────

def _slots(params, base):
    """[base, base2, base3] со strip'нутыми значениями."""
    return [
        (params.get(base)       or "").strip(),
        (params.get(base + "2") or "").strip(),
        (params.get(base + "3") or "").strip(),
    ]


def _warn_slots(base, slots):
    """Предупреждает о дырках и совпадающих кодах в слотах."""
    filled = [i for i, v in enumerate(slots) if v]
    if not filled:
        return

    lo, hi = min(filled), max(filled)
    for i in range(lo, hi + 1):
        if not slots[i]:
            print(f"[export_vars] warning: {base}{i+1} заполнен, "
                  f"но {base}{i} пуст", file=sys.stderr)

    seen = {}
    for i, v in enumerate(slots):
        if not v:
            continue
        if v in seen:
            print(f"[export_vars] warning: {base}{i+1} совпадает с "
                  f"{base}{seen[v]+1}", file=sys.stderr)
        else:
            seen[v] = i


def check_slots(params):
    """Диагностика по обоим наборам слотов."""
    _warn_slots("code_order", _slots(params, "code_order"))
    _warn_slots("code_hmi",   _slots(params, "code_hmi"))


# ─── Производные «плоские» переменные (обратная совместимость) ────────────────

def build_code_order_line(params):
    """Все коды ЮНИТ, по одному на строке. Порядок: 3 → 2 → 1."""
    vals = [v for v in reversed(_slots(params, "code_order")) if v]
    return "\n".join(vals)


def build_code_order_hmi_line(params):
    """Все коды ИЧМ, по одному на строке. Порядок: 3 → 2 → 1."""
    vals = [v for v in reversed(_slots(params, "code_hmi")) if v]
    return "\n".join(vals)


def add_derived(params, mk_lines, key, value):
    """Кладёт производную переменную и в params (JSON), и в mk_lines (.mk)."""
    params[key] = value
    mk_lines.append(f"{key.upper()} := {make_escape(_to_mk_value(value))}")


# ─── main ─────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) != 4:
        print("Usage: export_vars.py <id> <excel> <out.mk>", file=sys.stderr)
        sys.exit(2)

    target_id, excel_path, out_mk = sys.argv[1], sys.argv[2], sys.argv[3]
    out_json = out_mk.replace(".mk", ".json")

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.worksheets[0]
    headers = [c.value for c in ws[1]]

    if any(h is None for h in headers):
        print("[export_vars] warning: пустые ячейки в заголовках",
              file=sys.stderr)

    for row in ws.iter_rows(min_row=2, values_only=True):
        rec = dict(zip(headers, row))
        if str(rec.get("id", "")).strip() != target_id:
            continue

        params = {}
        mk_lines = []

        for key, value in rec.items():
            if key is None:
                continue
            name = str(key).strip().replace("\xa0", "").lower()

            if value is None:
                value = ""
            elif name == "date":
                value = fmt_date(value)
            elif name == "description":
                value = normalize_description(value)

            value = str(value)
            params[name] = value

            mk_lines.append(
                f"{name.upper()} := {make_escape(_to_mk_value(value))}"
            )

        # ── Производные «плоские» переменные ──
        # Оставлены на случай, если markdown-разделы ещё используют
        # {{ code_order_unit }} / {{ code_order_hmi }}.
        # Если нигде не нужны — можно удалить эти 4 строки.
        add_derived(params, mk_lines, "code_order_unit",
                    build_code_order_line(params))
        add_derived(params, mk_lines, "code_order_hmi",
                    build_code_order_hmi_line(params))

        # ── Диагностика заполнения слотов ──
        check_slots(params)

        with open(out_mk, "w", encoding="utf-8") as f:
            f.write("\n".join(mk_lines) + "\n")

        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(params, f, ensure_ascii=False, indent=2)

        return

    # ID не найден: файлы НЕ перезаписываем, чтобы не терять предыдущий результат
    print(f"ID '{target_id}' not found in {excel_path}", file=sys.stderr)
    sys.exit(1)


if __name__ == "__main__":
    main()