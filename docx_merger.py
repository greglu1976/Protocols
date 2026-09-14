import os
from docx import Document
from docxcompose.composer import Composer

def merge_docx_files(title_path, main_path, output_path):
    # 1. Загружаем титульный лист
    master = Document(title_path)
    composer = Composer(master)
    
    # 2. Загружаем основную часть (содержание и пункты)
    sub_doc = Document(main_path)
    
    # 3. Добавляем разрыв страницы, чтобы содержание началось со 2-й страницы
    #master.add_page_break()
    
    # 4. Склеиваем документы с сохранением всех стилей и таблиц МЭК 61850
    composer.append(sub_doc)
    
    # 5. Сохраняем готовый протокол
    composer.save(output_path)
    print(f"Готово! Файл сохранен как: {output_path}")

# Названия ваших файлов (замените при необходимости)
title_file = "title_page.docx"   
main_file = "main_content.docx"  
result_file = "protocol.docx" 

if os.path.exists(title_file) and os.path.exists(main_file):
    merge_docx_files(title_file, main_file, result_file)
else:
    print("Ошибка: Убедитесь, что файлы title_page.docx и main_content.docx лежат в одной папке со скриптом.")
