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
SECTIONS_JSON  = $(CABINET)/sections.json

MAIN_CONTENT = $(TEMP_DIR)/main_content.docx
COMBINED_MD  = $(TEMP_DIR)/combined.md

REFERENCE = templates/title_page.docx
FILTERS = --lua-filter=lua/pagebreak.lua --filter pandoc-crossref

PANDOC_OPTS = --standalone $(FILTERS) --reference-doc=$(REFERENCE)
#PANDOC_OPTS = --standalone --number-sections $(FILTERS) --reference-doc=$(REFERENCE) --lua-filter=lua/start-at-10.lua

POSTPROCESS        = python/postprocess.py
TABLE_PREPROCESSOR = python/preprocess.py
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
	@$(foreach c,$(CABINETS),$(MAKE) --no-print-directory ONE=1 CABINET=$(c) \
	    POSTPROCESS=$(POSTPROCESS) TABLE_PREPROCESSOR=$(TABLE_PREPROCESSOR) \
	    EXPAND_TABLES=$(EXPAND_TABLES) _one && ) \
	    $(MAKE) --no-print-directory clean-temp && \
	    echo All done.

# make — собрать один CABINET и подчистить TEMP
one:
	@$(MAKE) --no-print-directory ONE=1 CABINET=$(CABINET) \
	    POSTPROCESS=$(POSTPROCESS) TABLE_PREPROCESSOR=$(TABLE_PREPROCESSOR) \
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
	@$(MAKE) --no-print-directory ONE=1 CABINET=$(CABINET) _show-one

_show-one:
	@echo "CABINET     = $(CABINET)"
	@echo "CABINET_ID  = $(CABINET_ID)"

$(TEMP_DIR):
	-@if not exist "$(subst /,\,$(TEMP_DIR))" mkdir "$(subst /,\,$(TEMP_DIR))" 2>nul

# combined.md: препроцессор таблиц, затем простая конкатенация исходников
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

# Мерж титульника (шаблон как есть) и основного содержимого
$(FINAL_TARGET): $(TITLE_TEMPLATE) $(MAIN_CONTENT) python/docx_merger.py
	@echo "Merging documents..."
	$(PYTHON) python/docx_merger.py "$(TITLE_TEMPLATE)" "$(MAIN_CONTENT)" "$(FINAL_TARGET)"
	@echo "Cleaning up intermediate files..."
	-@if exist "$(subst /,\,$(MAIN_CONTENT))" del /Q "$(subst /,\,$(MAIN_CONTENT))"

# Убираем все финальные docx по каждому кабинету и чистим TEMP
clean:
	-@$(foreach c,$(CABINETS),if exist "$(subst /,\,$(c))\protocol_$(c).docx" del /Q "$(subst /,\,$(c))\protocol_$(c).docx" &) \
		if exist "$(subst /,\,$(TEMP_ROOT))" rmdir /S /Q "$(subst /,\,$(TEMP_ROOT))" 2>nul

# Только удаление TEMP
clean-temp:
	-@if exist "$(subst /,\,$(TEMP_ROOT))" rmdir /S /Q "$(subst /,\,$(TEMP_ROOT))" 2>nul