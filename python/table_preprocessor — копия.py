#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Заглушка препроцессора таблиц.
Находит теги <!-- INSERT_TABLE: ... --> в Markdown-файлах 
и заменяет их на ASCII-таблицы из Excel.
"""
import sys
import os
import re
from pathlib import Path

def process_file(md_path):
    """Обрабатывает один Markdown-файл."""
    content = Path(md_path).read_text(encoding='utf-8')
    
    # TODO: Реальная логика парсинга тегов и вставки таблиц
    # Пока просто выводим информацию о найденных тегах
    tags = re.findall(r'<!--\s*INSERT_TABLE:[^>]+-->', content)
    if tags:
        print(f"[PREPROCESS] Found {len(tags)} table tag(s) in {md_path}")
        for tag in tags[:3]:  # Показываем первые 3
            print(f"  - {tag.strip()}")
    else:
        print(f"[PREPROCESS] No table tags in {md_path}")
    
    # В реальном режиме здесь будет замена тегов на ASCII-сетки
    # Сейчас возвращаем контент без изменений
    return content

def main():
    if len(sys.argv) < 2:
        print("Usage: python table_preprocessor.py file1.md [file2.md ...]", file=sys.stderr)
        sys.exit(1)
    
    processed_count = 0
    for md_file in sys.argv[1:]:
        if not os.path.isfile(md_file):
            print(f"[PREPROCESS] Warning: File not found: {md_file}", file=sys.stderr)
            continue
        
        result = process_file(md_file)
        
        # Заглушка: перезаписываем файл тем же контентом
        # В реальности здесь была бы запись обработанного контента
        Path(md_file).write_text(result, encoding='utf-8')
        processed_count += 1
    
    print(f"[PREPROCESS] Done. Processed {processed_count} file(s).")

if __name__ == "__main__":
    main()