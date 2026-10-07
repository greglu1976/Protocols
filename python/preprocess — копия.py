#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Главный оркестратор препроцессинга Markdown-файлов.
Вызывает специализированные обработчики в заданном порядке.
"""
import sys
import os, re
from pathlib import Path
from table_generator import generate_table_ascii

# Регулярка для парных тегов (копируем из старого скрипта)
TAG_PARAMS_RE = re.compile(r'(?:file=(?P<file>[^,\s]+),?\s*)?(?:,\s*sheet(?:_name)?=(?P<sheet>[^,]+?))?')
OPEN_TAG = "<!-- INSERT_TABLE:"
CLOSE_TAG = "<!-- END_INSERT_TABLE -->"

def process_tables(content, md_file_dir):
    """Обрабатывает все блоки INSERT_TABLE в контенте."""
    result_parts = []
    pos = 0
    
    while pos < len(content):
        open_idx = content.find(OPEN_TAG, pos)
        if open_idx == -1:
            result_parts.append(content[pos:])
            break
            
        result_parts.append(content[pos:open_idx])
        open_end_idx = content.find('-->', open_idx)
        
        if open_end_idx == -1:
            result_parts.append(content[open_idx:])
            break
            
        tag_content = content[open_idx + len(OPEN_TAG):open_end_idx]
        params_match = TAG_PARAMS_RE.match(tag_content.strip())
        params = params_match.groupdict() if params_match else {}
        
        file_rel = params.get('file')
        close_idx = content.find(CLOSE_TAG, open_end_idx)
        
        if not file_rel or close_idx == -1:
            result_parts.append(content[open_idx:close_idx + len(CLOSE_TAG) if close_idx != -1 else open_end_idx + 3])
            pos = (close_idx + len(CLOSE_TAG)) if close_idx != -1 else (open_end_idx + 3)
            continue
            
        sheet_name = params.get('sheet') or 'Лист1'
        file_abs = os.path.normpath(os.path.join(md_file_dir, file_rel))
        
        try:
            new_table = generate_table_ascii(file_abs, sheet_name=sheet_name)
            table_block = f"\n{new_table}\n"
            print(f"[TABLE] Updated from '{file_rel}'")
        except Exception as e:
            table_block = f"\n[ERROR: {e}]\n"
            print(f"[TABLE] ERROR processing '{file_rel}': {e}", file=sys.stderr)
            
        full_open_tag = content[open_idx:open_end_idx + 3]
        result_parts.extend([full_open_tag, table_block, CLOSE_TAG])
        pos = close_idx + len(CLOSE_TAG)
        
    return ''.join(result_parts)

def run_pipeline(md_path):
    """Запускает полную цепочку препроцессинга для одного файла."""
    content = Path(md_path).read_text(encoding='utf-8')
    md_dir = os.path.dirname(os.path.abspath(md_path))
    
    # Этап 1: Таблицы
    new_content = process_tables(content, md_dir)
    # Этап 2: Здесь будут другие обработчики
    # new_content = process_images(new_content, md_dir)
    # new_content = process_calculations(new_content)
    
    if new_content != content:
        Path(md_path).write_text(new_content, encoding='utf-8')
        return True
    return False

def main():
    if len(sys.argv) < 2:
        print("Usage: python preprocess.py file1.md [file2.md ...]", file=sys.stderr)
        sys.exit(1)

    updated_count = 0
    for md_file in sys.argv[1:]:
        if not os.path.isfile(md_file):
            print(f"[PREPROCESS] Warning: File not found: {md_file}", file=sys.stderr)
            continue
        if run_pipeline(md_file):
            updated_count += 1
            
    print(f"[PREPROCESS] Done. Updated {updated_count}/{len(sys.argv)-1} file(s).")

if __name__ == "__main__":
    main()