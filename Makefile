# --- Переменные ---
CABINET    ?= SET_210.05-0
CABINET_ID ?= $(patsubst SET_%,%,$(CABINET))
EXCEL      ?= vars_parsing.xlsx

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

REFERENCE = templates/title_page.docx
FILTERS = --lua-filter=lua/pagebreak.lua --filter pandoc-crossref

PANDOC_OPTS = --standalone $(FILTERS) --reference-doc=$(REFERENCE)
#PANDOC_OPTS = --standalone --number-sections $(FILTERS) --reference-doc=$(REFERENCE) --lua-filter=lua/start-at-10.lua

CABINET_VARS = $(TEMP_DIR)/cabinet_vars.mk
CABINET_JSON = $(TEMP_DIR)/cabinet_vars.json

POSTPROCESS        = python/postprocess.py
TABLE_PREPROCESSOR = python/preprocess.py
EXPAND_MD_TABLES   = python/expand_md_tables.py

# --- Цели ---
.PHONY: all one _one show show-one _show-one list clean clean-temp

# make без аргументов — один CABINET
.DEFAULT_GOAL := one

# --- Правило генерации переменных кабинета (для явных зависимостей) ---
$(CABINET_VARS) $(CABINET_JSON): $(EXCEL) python/export_vars.py | $(TEMP_DIR)
	$(PYTHON) python/export_vars.py "$(CABINET_ID)" "$(EXCEL)" "$(CABINET_VARS)"

# --- Переменные кабинета грузятся только в подмейке ---
ifeq ($(ONE),1)

# Создаём TEMP и генерируем переменные до -include,
# чтобы NAME/CODE/… были доступны для проверки ниже.
$(shell if not exist "$(subst /,\,$(TEMP_DIR))" mkdir "$(subst /,\,$(TEMP_DIR))")
$(shell $(PYTHON) python/export_vars.py "$(CABINET_ID)" "$(EXCEL)" "$(CABINET_VARS)")
-include $(CABINET_VARS)

ifeq ($(strip $(NAME)),)
$(error No data for CABINET_ID='$(CABINET_ID)' (CABINET='$(CABINET)') in $(EXCEL). Check 'id' column.)
endif

SOURCE := $(shell $(PYTHON) -c "import json, os; files=json.load(open('$(SECTIONS_JSON)')); print(' '.join([os.path.join('$(CABINET)', f) if not os.path.isabs(f) else f for f in files]))")
endif

# make all — собрать каждый найденный кабинет последовательным вызовом подмейки
all:
	@$(foreach c,$(CABINETS),$(MAKE) --no-print-directory ONE=1 CABINET=$(c) \
	    POSTPROCESS=$(POSTPROCESS) TABLE_PREPROCESSOR=$(TABLE_PREPROCESSOR) \
	    EXPAND_MD_TABLES=$(EXPAND_MD_TABLES) _one && ) \
	    $(MAKE) --no-print-directory clean-temp && \
	    echo All done.

# make — собрать один CABINET и подчистить TEMP
one:
	@$(MAKE) --no-print-directory ONE=1 CABINET=$(CABINET) \
	    POSTPROCESS=$(POSTPROCESS) TABLE_PREPROCESSOR=$(TABLE_PREPROCESSOR) \
	    EXPAND_MD_TABLES=$(EXPAND_MD_TABLES) _one && \
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
	@echo "EXCEL           = $(EXCEL)"

show-one:
	@$(MAKE) --no-print-directory ONE=1 CABINET=$(CABINET) _show-one

_show-one:
	@echo "CABINET     = $(CABINET)"
	@echo "CABINET_ID  = $(CABINET_ID)"
	@echo "EXCEL       = $(EXCEL)"
	@echo "NAME        = $(NAME)"
	@echo "CODE        = $(CODE)"
	@echo "VERSION     = $(VERSION)"
	@echo "DATE        = $(DATE)"
	@echo "MANUAL      = $(MANUAL)"
	@echo "DESCRIPTION = $(DESCRIPTION)"

$(TEMP_DIR):
	-@if not exist "$(subst /,\,$(TEMP_DIR))" mkdir "$(subst /,\,$(TEMP_DIR))" 2>nul

# Титульник: только подстановка {{ vars }} — без размножения таблиц
$(TITLE_FILLED): $(TITLE_TEMPLATE) $(CABINET_JSON) python/placeholder_filler.py | $(TEMP_DIR)
	@echo "Rendering title page for $(CABINET)..."
	$(PYTHON) python/placeholder_filler.py \
		--template "$(TITLE_TEMPLATE)" \
		--params   "$(CABINET_JSON)" \
		--out      "$(TITLE_FILLED)"

# combined.md: препроцессор таблиц → подстановка {{ vars }} → размножение <<REPEAT_*>> / <<UNITS_*>>
$(COMBINED_MD): $(SOURCE) $(CABINET_JSON) $(SECTIONS_JSON) \
                $(TABLE_PREPROCESSOR) python/md_subst.py $(EXPAND_MD_TABLES) | $(TEMP_DIR)
	@echo "Preprocessing tables in source files..."
	$(PYTHON) $(TABLE_PREPROCESSOR) $(SOURCE)
	@echo "Substituting {{ vars }} in markdown..."
	$(PYTHON) python/md_subst.py "$(CABINET_JSON)" "$(COMBINED_MD)" $(SOURCE)
	@echo "Expanding UNIT/HMI/PRV tables in markdown..."
	$(PYTHON) $(EXPAND_MD_TABLES) "$(COMBINED_MD)" "$(CABINET_JSON)"

# main_content.docx через Pandoc
$(MAIN_CONTENT): $(COMBINED_MD) $(REFERENCE) lua/pagebreak.lua
	@echo "Generating main content via Pandoc..."
	@echo "Using sections: $(SOURCE)"
	$(PANDOC) "$(COMBINED_MD)" -o "$(MAIN_CONTENT)" $(PANDOC_OPTS)

$(FINAL_TARGET): $(TITLE_FILLED) $(MAIN_CONTENT) python/docx_merger.py
	@echo "Merging documents..."
	$(PYTHON) python/docx_merger.py "$(TITLE_FILLED)" "$(MAIN_CONTENT)" "$(FINAL_TARGET)"
	@echo "Cleaning up intermediate files..."
	-@if exist "$(subst /,\,$(MAIN_CONTENT))" del /Q "$(subst /,\,$(MAIN_CONTENT))"
	-@if exist "$(subst /,\,$(TITLE_FILLED))" del /Q "$(subst /,\,$(TITLE_FILLED))"

# Убираем все финальные docx по каждому кабинету и чистим TEMP
clean:
	-@$(foreach c,$(CABINETS),if exist "$(subst /,\,$(c))\protocol_$(c).docx" del /Q "$(subst /,\,$(c))\protocol_$(c).docx" &) \
		if exist "$(subst /,\,$(TEMP_ROOT))" rmdir /S /Q "$(subst /,\,$(TEMP_ROOT))" 2>nul

# Только удаление TEMP
clean-temp:
	-@if exist "$(subst /,\,$(TEMP_ROOT))" rmdir /S /Q "$(subst /,\,$(TEMP_ROOT))" 2>nul