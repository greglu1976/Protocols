# --- Переменные ---
CABINET ?= ШЭТ_440.01-0

# Папка для временных файлов
TEMP_DIR = TEMP

# Имя готового файла в папке кабинета
FINAL_TARGET = $(CABINET)/protocol_$(CABINET).docx

TITLE_TEMPLATE = templates/title_page.docx
TITLE_FILLED   = $(TEMP_DIR)/title_page_filled.docx
PARAMS_JSON    = $(CABINET)/params.json

MAIN_CONTENT = $(TEMP_DIR)/main_content.docx

SOURCE = $(CABINET)/section1.md $(CABINET)/section2.md $(CABINET)/section3.md \
         $(CABINET)/section4.md $(CABINET)/section5.md $(CABINET)/section6.md \
         $(CABINET)/section7.md

REFERENCE = templates/title_page.docx
FILTERS = --lua-filter=lua/pagebreak.lua --filter pandoc-crossref

PANDOC ?= pandoc
PYTHON ?= python

PANDOC_OPTS = --standalone --number-sections $(FILTERS) --reference-doc=$(REFERENCE)

# --- Цели (Targets) ---
.PHONY: all clean

all: $(FINAL_TARGET)

# Создаём папку TEMP если её нет
$(TEMP_DIR):
	@if not exist "$(TEMP_DIR)" mkdir "$(TEMP_DIR)"

# 1. Заполнение титульного листа (docxtpl)
$(TITLE_FILLED): $(TITLE_TEMPLATE) $(PARAMS_JSON) python/placeholder_filler.py | $(TEMP_DIR)
	@echo "Rendering title page..."
	@chcp 65001 >nul 2>&1 || true
	$(PYTHON) python/placeholder_filler.py "$(TITLE_TEMPLATE)" "$(CABINET)" "$(TITLE_FILLED)"

# 2. Основная часть через Pandoc
$(MAIN_CONTENT): $(SOURCE) $(REFERENCE) lua/pagebreak.lua | $(TEMP_DIR)
	@echo "Generating main content via Pandoc..."
	@chcp 65001 >nul 2>&1 || true
	$(PANDOC) $(SOURCE) -o $(MAIN_CONTENT) $(PANDOC_OPTS)

# 3. Слияние титульника и основной части
$(FINAL_TARGET): $(TITLE_FILLED) $(MAIN_CONTENT) python/docx_merger.py
	@echo "Merging documents..."
	@chcp 65001 >nul 2>&1 || true
	$(PYTHON) python/docx_merger.py "$(TITLE_FILLED)" "$(MAIN_CONTENT)" "$(FINAL_TARGET)"
	@echo "Cleaning up intermediate files..."
	@if exist "$(MAIN_CONTENT)" del /Q "$(MAIN_CONTENT)"
	@if exist "$(TITLE_FILLED)" del /Q "$(TITLE_FILLED)"

# Очистка
clean:
	@chcp 65001 >nul 2>&1 || true
	@if exist "$(TEMP_DIR)\$(notdir $(MAIN_CONTENT))" del /Q "$(TEMP_DIR)\$(notdir $(MAIN_CONTENT))"
	@if exist "$(TEMP_DIR)\$(notdir $(TITLE_FILLED))" del /Q "$(TEMP_DIR)\$(notdir $(TITLE_FILLED))"
	@for %%f in ($(CABINET)\protocol_*.docx) do @if exist "%%f" del /Q "%%f"