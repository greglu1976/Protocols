#!/usr/bin/env python
# -*- coding: utf-8 -*-
import json
import sys
from pathlib import Path
from docxtpl import DocxTemplate

def fill_title_page(template_path: Path, params: dict, output_path: Path) -> None:
    doc = DocxTemplate(str(template_path))
    doc.render(params)  # params = {"object_name": "...", "customer": "...", ...}
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))

def main(argv: list) -> int:
    if len(argv) != 4:
        print("Usage: title_page.py <template.docx> <cabinet_dir> <output.docx>")
        return 2

    template_path = Path(argv[1])
    cabinet_dir = Path(argv[2])
    output_path = Path(argv[3])

    params_file = cabinet_dir / "params.json"
    if not params_file.exists():
        print(f"Ошибка: {params_file} не найден", file=sys.stderr)
        return 1

    with params_file.open("r", encoding="utf-8") as f:
        params = json.load(f)

    fill_title_page(template_path, params, output_path)
    print(f"Титульный лист сохранён: {output_path}")
    return 0

if __name__ == "__main__":
    sys.exit(main(sys.argv))