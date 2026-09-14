import sys
import os
from docx import Document
from docxcompose.composer import Composer

def merge_docx_files(title_path, main_path, output_path):
    # 1. Загружаем титульный лист
    master = Document(title_path)
    composer = Composer(master)
    
    # 2. Загружаем основную часть (содержание и пункты)
    sub_doc = Document(main_path)
    
    # 3. Склеиваем документы с сохранением всех стилей
    composer.append(sub_doc)
    
    # 4. Сохраняем готовый протокол
    composer.save(output_path)
    print(f"Готово! Файл сохранен как: " + output_path)

# Проверяем, переданы ли пути к файлам из Makefile (нужно ровно 3 аргумента)
if len(sys.argv) == 4:
    title_file = sys.argv[1]
    main_file = sys.argv[2]
    result_file = sys.argv[3]
    
    if os.path.exists(title_file) and os.path.exists(main_file):
        merge_docx_files(title_file, main_file, result_file)
    else:
        print(f"Ошибка: Файлы не найдены. Искали: '{title_file}' и '{main_file}'")
else:
    print("Ошибка: Скрипту нужно передать 3 параметра (title.docx main.docx final.docx)")