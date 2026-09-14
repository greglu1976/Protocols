# --- Переменные ---
FINAL_TARGET = final_protocol.docx
TITLE_PAGE = title_page.docx
MAIN_CONTENT = main_content.docx

SOURCE = section1.md section2.md
REFERENCE = reference.docx
FILTERS = --lua-filter=pagebreak.lua --filter pandoc-crossref

PANDOC ?= pandoc
PYTHON ?= python

# Опции Pandoc
PANDOC_OPTS = --standalone --number-sections $(FILTERS) --reference-doc=$(REFERENCE)

# --- Цели (Targets) ---
all: $(FINAL_TARGET)

# 1. Основное правило: собираем финальный файл из титульника и сгенерированного контента
$(FINAL_TARGET): $(MAIN_CONTENT) $(TITLE_PAGE) docx_merger.py
	@echo "Объединяю титульный лист и основную часть с помощью Python..."
	$(PYTHON) docx_merger.py

# 2. Правило для сборки основной части через Pandoc (зависит от md-файлов)
$(MAIN_CONTENT): $(SOURCE) $(REFERENCE) pagebreak.lua
	@echo "Генерирую основную часть из Markdown через Pandoc..."
	$(PANDOC) $(SOURCE) -o $(MAIN_CONTENT) $(PANDOC_OPTS)

# Очистка всех сгенерированных файлов
clean:
	rm -f $(MAIN_CONTENT) $(FINAL_TARGET)

.PHONY: all clean
