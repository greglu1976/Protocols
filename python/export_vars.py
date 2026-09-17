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
    s = s.replace("\\n", "\n")   # литерал -> реальный newline
    s = s.replace("\r\n", "\n")  # CRLF -> LF
    s = s.replace("\r", "\n")    # одиночный CR -> LF
    # убираем пробелы в конце/начале каждой строки
    s = "\n".join(part.strip() for part in s.split("\n"))
    return s


def make_escape(s):
    return s.replace("\\", "\\\\").replace("$", "$$").replace("#", "\\#")


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

            # В .mk переводы строк заменяем на пробел (иначе make сломается)
            mk_value = value.replace("\n", " ")
            mk_lines.append(f"{name.upper()} := {make_escape(mk_value)}")

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