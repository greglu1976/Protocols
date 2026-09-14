# --- Переменные ---
CABINET ?= ШЭТ_440.01-0

# Имя готового файла: protocol_ШЭТ_440.01-0.docx
FINAL_TARGET = protocol_$(CABINET).docx
TITLE_PAGE = templates/title_page.docx
MAIN_CONTENT = main_content.docx

# Кавычки не нужны, так как пробелов нет
SOURCE = $(CABINET)/section1.md $(CABINET)/section2.md
REFERENCE = templates/reference.docx
FILTERS = --lua-filter=lua/pagebreak.lua --filter pandoc-crossref

PANDOC ?= pandoc
PYTHON ?= python

PANDOC_OPTS = --standalone --number-sections $(FILTERS) --reference-doc=$(REFERENCE)

# --- Цели (Targets) ---
all: $(FINAL_TARGET)

# 1. Основное правило (Python)
 $(FINAL_TARGET): $(MAIN_CONTENT) $(TITLE_PAGE) python/docx_merger.py
	@echo "Merging documents..."
	$(PYTHON) python/docx_merger.py $(TITLE_PAGE) $(MAIN_CONTENT) $(FINAL_TARGET)
	@echo "Cleaning up intermediate file..."
	-del /Q $(MAIN_CONTENT)

# 2. Правило для сборки основной части (Pandoc)
 $(MAIN_CONTENT): $(SOURCE) $(REFERENCE) lua/pagebreak.lua
	@echo "Generating main content via Pandoc..."
	$(PANDOC) $(SOURCE) -o $(MAIN_CONTENT) $(PANDOC_OPTS)

# Очистка
clean:
	rm -f $(MAIN_CONTENT) protocol_*.docx