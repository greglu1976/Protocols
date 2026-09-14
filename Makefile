# --- Переменные ---
CABINET = 440.01-0

FINAL_TARGET = final_protocol.docx
TITLE_PAGE = templates/title_page.docx
MAIN_CONTENT = main_content.docx

SOURCE = $(CABINET)/section1.md $(CABINET)/section2.md
REFERENCE = templates/reference.docx
FILTERS = --lua-filter=lua/pagebreak.lua --filter pandoc-crossref

PANDOC ?= pandoc
PYTHON ?= python

# Опции Pandoc
PANDOC_OPTS = --standalone --number-sections $(FILTERS) --reference-doc=$(REFERENCE)

# --- Цели (Targets) ---
all: $(FINAL_TARGET)

# 1. Основное правило: собираем финальный файл
$(FINAL_TARGET): $(MAIN_CONTENT) $(TITLE_PAGE) python/docx_merger.py
	@echo "Merging documents..."
	$(PYTHON) python/docx_merger.py $(TITLE_PAGE) $(MAIN_CONTENT) $(FINAL_TARGET)

# 2. Правило для сборки основной части через Pandoc (зависит от md-файлов)
$(MAIN_CONTENT): $(SOURCE) $(REFERENCE) lua/pagebreak.lua
	@echo "Генерирую основную часть из Markdown через Pandoc..."
	$(PANDOC) $(SOURCE) -o $(MAIN_CONTENT) $(PANDOC_OPTS)

# Очистка всех сгенерированных файлов
clean:
	rm -f $(MAIN_CONTENT) $(FINAL_TARGET)

.PHONY: all clean
