# --- Переменные ---
CABINET ?= SET_440.01-0

TEMP_DIR = TEMP

FINAL_TARGET = $(CABINET)/protocol_$(CABINET).docx

TITLE_TEMPLATE = templates/title_page.docx
TITLE_FILLED   = $(TEMP_DIR)/title_page_filled.docx
PARAMS_JSON    = $(CABINET)/params.json
SECTIONS_JSON  = $(CABINET)/sections.json

MAIN_CONTENT = $(TEMP_DIR)/main_content.docx

REFERENCE = templates/title_page.docx
FILTERS = --lua-filter=lua/pagebreak.lua --filter pandoc-crossref

PANDOC ?= pandoc
PYTHON ?= python

PANDOC_OPTS = --standalone --number-sections $(FILTERS) --reference-doc=$(REFERENCE)

# Динамическая генерация SOURCE из sections.json
# Поддерживает относительные пути (добавляет CABINET/) и абсолютные пути
# SOURCE = $(shell python -c "import json, os; files=json.load(open('$(SECTIONS_JSON)')); print(' '.join([os.path.join('$(CABINET)', f) if not os.path.isabs(f) else f for f in files]))")
SOURCE = $(shell python -c "import json, os; files=json.load(open('$(SECTIONS_JSON)')); print(' '.join([os.path.join('$(CABINET)', f) if not os.path.isabs(f) else f for f in files]))")

# --- Цели (Targets) ---
.PHONY: all clean

all: $(FINAL_TARGET)

$(TEMP_DIR):
	-@if not exist "$(TEMP_DIR)" mkdir "$(TEMP_DIR)" 2>nul

$(TITLE_FILLED): $(TITLE_TEMPLATE) $(PARAMS_JSON) python/placeholder_filler.py | $(TEMP_DIR)
	@echo "Rendering title page..."
	$(PYTHON) python/placeholder_filler.py "$(TITLE_TEMPLATE)" "$(CABINET)" "$(TITLE_FILLED)"

$(MAIN_CONTENT): $(SOURCE) $(REFERENCE) lua/pagebreak.lua $(SECTIONS_JSON) | $(TEMP_DIR)
	@echo "Generating main content via Pandoc..."
	@echo "Using sections: $(SOURCE)"
	$(PANDOC) $(SOURCE) -o "$(MAIN_CONTENT)" $(PANDOC_OPTS)

$(FINAL_TARGET): $(TITLE_FILLED) $(MAIN_CONTENT) python/docx_merger.py
	@echo "Merging documents..."
	$(PYTHON) python/docx_merger.py "$(TITLE_FILLED)" "$(MAIN_CONTENT)" "$(FINAL_TARGET)"
	@echo "Cleaning up intermediate files..."
	-@if exist "$(subst /,\,$(MAIN_CONTENT))" del /Q "$(subst /,\,$(MAIN_CONTENT))"
	-@if exist "$(subst /,\,$(TITLE_FILLED))" del /Q "$(subst /,\,$(TITLE_FILLED))"

clean:
	-@if exist "$(subst /,\,$(MAIN_CONTENT))" del /Q "$(subst /,\,$(MAIN_CONTENT))"
	-@if exist "$(subst /,\,$(TITLE_FILLED))" del /Q "$(subst /,\,$(TITLE_FILLED))"
	-@if exist "$(subst /,\,$(FINAL_TARGET))" del /Q "$(subst /,\,$(FINAL_TARGET))"
	-@if exist "$(subst /,\,$(TEMP_DIR))" rmdir /Q "$(subst /,\,$(TEMP_DIR))" 2>nul