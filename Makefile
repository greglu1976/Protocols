# --- Переменные ---
CABINET    ?= SET_240.01-0
CABINET_ID ?= $(patsubst SET_%,%,$(CABINET))

PANDOC ?= pandoc
PYTHON ?= python

# ─── UTF-8 для всех дочерних Python ───
export PYTHONIOENCODING := utf-8
export PYTHONUTF8 := 1

# Все кабинеты: папки SET_*, в которых есть sections.json
CABINETS := $(shell $(PYTHON) -c "import glob,os; print(' '.join(sorted(os.path.dirname(p).replace(os.sep,'/') for p in glob.glob('SET_*/sections.json'))))")
ifeq ($(strip $(CABINETS)),)
CABINETS := $(CABINET)
endif

TEMP_ROOT = TEMP
TEMP_DIR  = $(TEMP_ROOT)/$(CABINET)

FINAL_TARGET = $(CABINET)/protocol_$(CABINET).docx

TITLE_TEMPLATE = templates/title_page.docx
TITLE_FILLED   = $(TEMP_DIR)/title_page_filled.docx
SECTIONS_JSON  = $(CABINET)/sections.json

MAIN_CONTENT = $(TEMP_DIR)/main_content.docx
COMBINED_MD  = $(TEMP_DIR)/combined.md
VARS_JSON    = $(TEMP_DIR)/vars.json

REFERENCE = templates/title_page.docx
FILTERS = --lua-filter=lua/pagebreak.lua --filter pandoc-crossref

PANDOC_OPTS = --standalone $(FILTERS) --reference-doc=$(REFERENCE)
#PANDOC_OPTS = --standalone --number-sections $(FILTERS) --reference-doc=$(REFERENCE) --lua-filter=lua/start-at-10.lua

POSTPROCESS        = python/postprocess.py
GEN_TECH_TABLES    = python/gen_tech_tables.py
TABLE_PREPROCESSOR = python/preprocess.py
FILL_TITLE         = python/fill_title.py
EXPAND_TABLES      = python/expand_docx_tables.py

# --- Цели ---
.PHONY: all one _one show show-one _show-one list clean clean-temp

# make без аргументов — один CABINET
.DEFAULT_GOAL := one

# --- Переменные кабинета грузятся только в подмейке ---
ifeq ($(ONE),1)

# Получаем список исходных файлов из sections.json
SOURCE := $(shell $(PYTHON) -c "import json, os; files=json.load(open('$(SECTIONS_JSON)')); print(' '.join([os.path.join('$(CABINET)', f) if not os.path.isabs(f) else f for f in files]))")

endif

# make all — собрать каждый найденный кабинет последовательным вызовом подмейки
all:
	@$(foreach c,$(CABINETS),$(MAKE) --no-print-directory ONE=1 CABINET=$(c) CABINET_ID=$(patsubst SET_%,%,$(c)) \
	    POSTPROCESS=$(POSTPROCESS) GEN_TECH_TABLES=$(GEN_TECH_TABLES) \
	    TABLE_PREPROCESSOR=$(TABLE_PREPROCESSOR) FILL_TITLE=$(FILL_TITLE) \
	    EXPAND_TABLES=$(EXPAND_TABLES) _one && ) \
	    $(MAKE) --no-print-directory clean-temp && \
	    echo All done.

one:
	@$(MAKE) --no-print-directory ONE=1 CABINET=$(CABINET) CABINET_ID=$(CABINET_ID) \
	    POSTPROCESS=$(POSTPROCESS) GEN_TECH_TABLES=$(GEN_TECH_TABLES) \
	    TABLE_PREPROCESSOR=$(TABLE_PREPROCESSOR) FILL_TITLE=$(FILL_TITLE) \
	    EXPAND_TABLES=$(EXPAND_TABLES) _one && \
	    $(MAKE) --no-print-directory clean-temp

# Сборка одного кабинета + постобработка
_one: $(FINAL_TARGET)
ifeq ($(TEST),1)
	@echo "[TEST MODE] Skipping post-processing for $(CABINET)..."
else
	@echo "Post-processing $(FINAL_TARGET) for $(CABINET)..."
	set TEST=$(TEST)&& $(PYTHON) $(POSTPROCESS) "$(FINAL_TARGET)"
endif

list:
	@echo "CABINETS = [$(CABINETS)]"
	@echo "Count    = $(words $(CABINETS))"

show:
	@echo "Default CABINET = $(CABINET)"
	@echo "All CABINETS    = $(CABINETS)"

show-one:
	@$(MAKE) --no-print-directory ONE=1 CABINET=$(CABINET) CABINET_ID=$(CABINET_ID) _show-one

_show-one:
	@echo "CABINET     = $(CABINET)"
	@echo "CABINET_ID  = $(CABINET_ID)"

$(TEMP_DIR):
	-@if not exist "$(subst /,\,$(TEMP_DIR))" mkdir "$(subst /,\,$(TEMP_DIR))" 2>nul

# vars.json из vars_parsing.xlsx
$(VARS_JSON): $(GEN_TECH_TABLES) vars_parsing.xlsx | $(TEMP_DIR)
	@echo "Generating vars.json for CABINET_ID=$(CABINET_ID)..."
	$(PYTHON) $(GEN_TECH_TABLES) $(CABINET_ID) > "$(VARS_JSON)"

# Заполненный титульник
$(TITLE_FILLED): $(TITLE_TEMPLATE) $(VARS_JSON) $(FILL_TITLE)
	@echo "Rendering title page for CABINET_ID=$(CABINET_ID)..."
	$(PYTHON) $(FILL_TITLE) $(CABINET_ID) "$(VARS_JSON)"

# combined.md: препроцессор таблиц, затем конкатенация исходников
$(COMBINED_MD): $(SOURCE) $(SECTIONS_JSON) $(TABLE_PREPROCESSOR) | $(TEMP_DIR)
	@echo "Preprocessing tables in source files..."
	$(PYTHON) $(TABLE_PREPROCESSOR) $(SOURCE)
	@echo "Combining markdown..."
	$(PYTHON) -c "import sys; out=open(sys.argv[1],'w',encoding='utf-8'); [out.write(open(p,encoding='utf-8').read()+'\n\n') for p in sys.argv[2:]]; out.close()" "$(COMBINED_MD)" $(SOURCE)

# main_content.docx через Pandoc
$(MAIN_CONTENT): $(COMBINED_MD) $(REFERENCE) lua/pagebreak.lua
	@echo "Generating main content via Pandoc..."
	@echo "Using sections: $(SOURCE)"
	$(PANDOC) "$(COMBINED_MD)" -o "$(MAIN_CONTENT)" $(PANDOC_OPTS)

# Мерж заполненного титульника и основного содержимого
$(FINAL_TARGET): $(TITLE_FILLED) $(MAIN_CONTENT) python/docx_merger.py
	@echo "Merging documents..."
	$(PYTHON) python/docx_merger.py "$(TITLE_FILLED)" "$(MAIN_CONTENT)" "$(FINAL_TARGET)"
	@echo "Cleaning up intermediate files..."
	-@if exist "$(subst /,\,$(MAIN_CONTENT))" del /Q "$(subst /,\,$(MAIN_CONTENT))"

# Убираем все финальные docx по каждому кабинету и чистим TEMP
clean:
	-@$(foreach c,$(CABINETS),if exist "$(subst /,\,$(c))\protocol_$(c).docx" del /Q "$(subst /,\,$(c))\protocol_$(c).docx" &) \
		if exist "$(subst /,\,$(TEMP_ROOT))" rmdir /S /Q "$(subst /,\,$(TEMP_ROOT))" 2>nul

# Только удаление TEMP
clean-temp:
	-@if exist "$(subst /,\,$(TEMP_ROOT))" rmdir /S /Q "$(subst /,\,$(TEMP_ROOT))" 2>nul