#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Заменяет стиль всех параграфов с 'Обычный' на 'Обычный1'.

Использование:
    python postprocess.py путь/к/protocol_XXX.docx
"""

import sys
import os
from docx import Document


TARGET_STYLE_NAME = "Основной текст с отступом 31"


def main():
    if len(sys.argv) < 2:
        print("Usage: python postprocess.py <docx_path>", file=sys.stderr)
        sys.exit(1)

    docx_path = sys.argv[1]

    if not os.path.isfile(docx_path):
        print(f"File not found: {docx_path}", file=sys.stderr)
        sys.exit(1)

    print(f"Post-processing: {docx_path}")

    doc = Document(docx_path)

    # проверим, что целевой стиль вообще есть в документе
    style_names = {s.name for s in doc.styles}
    if TARGET_STYLE_NAME not in style_names:
        print(
            f"Style '{TARGET_STYLE_NAME}' not found in document. "
            f"Available styles: {sorted(style_names)}",
            file=sys.stderr,
        )
        sys.exit(2)

    replaced = 0
    for p in doc.paragraphs:
        # style может быть None у пустых/служебных параграфов
        if p.style is None:
            continue
        # 'Обычный' в русском Word имеет внутреннее имя 'Normal'
        if p.style.name in ("Обычный", "Normal"):
            p.style = doc.styles[TARGET_STYLE_NAME]
            replaced += 1

    print(f"Replaced paragraphs: {replaced}")

    doc.save(docx_path)
    print("Saved:", docx_path)
    print("Done.")


if __name__ == "__main__":
    main()