# -*- coding: utf-8 -*-
import json
import sys
from datetime import datetime
import openpyxl


def fmt_date(value):
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    s = str(value).strip()
    if len(s) >= 10 and s[4] == "-" and s[7] == "-":
        y, m, d = s[:10].split("-")
        return f"{d}.{m}.{y}"
    return s


def normalize_description(text):
    """Литерал \n (обратный слэш + n) превращаем в настоящий перевод строки."""
    s = str(text)
    s = s.replace("\\n", "\n")
    s = s.replace("\r\n", "\n")
    s = s.replace("\r", "\n")
    s = "\n".join(part.strip() for part in s.split("\n"))
    return s


def make_escape(s):
    return s.replace("\\", "\\\\").replace("$", "$$").replace("#", "\\#")


def _two_codes(params, key1, key2):
    """Возвращает (code1, code2) — обе строки, strip'нутые, '' если пусто."""
    return (
        (params.get(key1) or "").strip(),
        (params.get(key2) or "").strip(),
    )


def _build_code_order_block(code1, code2, single_label):
    """
    Собирает блок «коды заказа» для markdown.
    single_label — что написать, если код только один
    (например: "Код заказа устройства" / "Код заказа ИЧМ").
    """
    if code1 and code2:
        return (
            "Коды заказа устройств:\n\n"
            f"{code1};\n\n"
            f"{code2}."
        )
    if code1:
        return f"{single_label}: {code1}"
    return ""


def build_code_order_line(params):
    """Блок кодов заказа для ЮНИТ (code_order / code_order2)."""
    c1, c2 = _two_codes(params, "code_order", "code_order2")
    return _build_code_order_block(c1, c2, "Код заказа устройства")


def build_code_order_hmi_line(params):
    """Блок кодов заказа для ИЧМ (code_hmi / code_hmi2)."""
    c1, c2 = _two_codes(params, "code_hmi", "code_hmi2")
    return _build_code_order_block(c1, c2, "Код заказа ИЧМ")


# ── Строки таблиц устройств ──

# ЮНИТ — 7 колонок: SN | Изготовитель | Год | Uпит | I1 | I2 | I3
_ROW_UNIT = "|                 | ООО Юнител Инжиниринг |             | =/~220  |             |             |             |"

# ИЧМ — 4 колонки: SN | Изготовитель | Год | Uпит
_ROW_HMI = "|                 | ООО Юнител Инжиниринг |             | =/~220  |"


def build_device_row2(params):
    """Вторая строка таблицы ЮНИТ. Только если заданы оба кода ЮНИТ."""
    c1, c2 = _two_codes(params, "code_order", "code_order2")
    return _ROW_UNIT if (c1 and c2) else ""


def build_device_row_hmi2(params):
    """Вторая строка таблицы ИЧМ. Только если заданы оба кода ИЧМ."""
    c1, c2 = _two_codes(params, "code_hmi", "code_hmi2")
    return _ROW_HMI if (c1 and c2) else ""


def add_derived(params, mk_lines, key, value):
    """Кладёт производную переменную и в params (JSON), и в mk_lines (.mk)."""
    params[key] = value
    mk_value = value.replace("\n", " ")
    mk_lines.append(f"{key.upper()} := {make_escape(mk_value)}")


def main():
    if len(sys.argv) != 4:
        print("Usage: export_vars.py <id> <excel> <out.mk>", file=sys.stderr)
        sys.exit(2)

    target_id, excel_path, out_mk = sys.argv[1], sys.argv[2], sys.argv[3]
    out_json = out_mk.replace(".mk", ".json")

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    headers = [c.value for c in ws[1]]

    for row in ws.iter_rows(min_row=2, values_only=True):
        rec = dict(zip(headers, row))
        if str(rec.get("id", "")).strip() != target_id:
            continue

        params = {}
        mk_lines = []

        for key, value in rec.items():
            if key is None:
                continue
            name = str(key).strip().lower()

            if value is None:
                value = ""
            elif name == "date":
                value = fmt_date(value)
            elif name == "description":
                value = normalize_description(value)

            value = str(value)
            params[name] = value

            mk_value = value.replace("\n", " ")
            mk_lines.append(f"{name.upper()} := {make_escape(mk_value)}")

        # ── Производные переменные: ЮНИТ ──
        add_derived(params, mk_lines, "code_order_unit",
                    build_code_order_line(params))
        add_derived(params, mk_lines, "device_row2",
                    build_device_row2(params))

        # ── Производные переменные: ИЧМ ──
        add_derived(params, mk_lines, "code_order_hmi",
                    build_code_order_hmi_line(params))
        add_derived(params, mk_lines, "device_row_hmi2",
                    build_device_row_hmi2(params))

        with open(out_mk, "w", encoding="utf-8") as f:
            f.write("\n".join(mk_lines) + "\n")

        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(params, f, ensure_ascii=False, indent=2)

        return

    open(out_mk, "w", encoding="utf-8").close()
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump({}, f)
    print(f"ID '{target_id}' not found in {excel_path}", file=sys.stderr)
    sys.exit(1)


if __name__ == "__main__":
    main()