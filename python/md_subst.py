# -*- coding: utf-8 -*-
"""Подставляет {{ key }} в markdown-файлах значениями из JSON.

Особенность: если строка содержит только плейсхолдер (например, {{ device_row2 }})
и после подстановки становится пустой — строка удаляется целиком. Это позволяет
опциональным строкам таблиц исчезать без пустых пропусков.

Использование:
    python md_subst.py <vars.json> <out.md> <in1.md> [in2.md ...]
"""
import json
import os
import re
import sys

PLACEHOLDER = re.compile(r"\{\{\s*([\w_]+)\s*\}\}")


def substitute(text, data, vars_path=""):
    def repl(m):
        key = m.group(1)
        if key not in data:
            print(f"[md_subst] warning: key '{key}' not found in {vars_path}",
                  file=sys.stderr)
            return ""
        return str(data[key])

    out = []
    for line in text.split("\n"):
        had_placeholder = PLACEHOLDER.search(line) is not None
        new_line = PLACEHOLDER.sub(repl, line)
        # Пустая строка, оставшаяся от плейсхолдера, выбрасывается целиком.
        if had_placeholder and new_line.strip() == "":
            continue
        out.append(new_line)
    return "\n".join(out)


def main():
    if len(sys.argv) < 4:
        print("Usage: md_subst.py <vars.json> <out.md> <in1.md> [in2.md ...]",
              file=sys.stderr)
        sys.exit(2)

    vars_path, out_path = sys.argv[1], sys.argv[2]
    inputs = sys.argv[3:]

    with open(vars_path, encoding="utf-8") as f:
        data = json.load(f)

    chunks = []
    for path in inputs:
        with open(path, encoding="utf-8") as f:
            text = f.read()
        text = substitute(text, data, vars_path)
        chunks.append(text)

    combined = "\n\n".join(chunks) + "\n"

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(combined)


if __name__ == "__main__":
    main()