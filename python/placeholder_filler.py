#!/usr/bin/env python
# -*- coding: utf-8 -*-
import argparse
import json
import sys
from pathlib import Path

from docxtpl import DocxTemplate, RichText


def fill_title_page(template_path: Path, params: dict, output_path: Path) -> None:
    # Оборачиваем поля с \n в RichText — иначе docxtpl выведет их в одну строку
    if "description" in params and params["description"]:
        params["description"] = RichText(str(params["description"]))

    doc = DocxTemplate(str(template_path))
    doc.render(params)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))


def main(argv: list) -> int:
    ap = argparse.ArgumentParser(description="Заполнение титульного листа")
    ap.add_argument("--template", required=True, help="шаблон .docx")
    ap.add_argument("--params",   required=True, help="JSON с переменными")
    ap.add_argument("--out",      required=True, help="куда сохранить .docx")
    args = ap.parse_args(argv[1:])

    params_path = Path(args.params)
    if not params_path.exists():
        print(f"Ошибка: {params_path} не найден", file=sys.stderr)
        return 1

    with params_path.open(encoding="utf-8") as f:
        params = json.load(f)

    fill_title_page(Path(args.template), params, Path(args.out))
    print(f"Титульный лист сохранён: {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))