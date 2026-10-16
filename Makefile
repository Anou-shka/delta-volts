# Rebuild everything from primary sources. `make all` runs every stage in order.
PY ?= python3

.PHONY: all data fetch extract charts model preview test clean

all: data charts model

data: fetch extract

fetch:            ## download filings into data/raw/ (skips files already present)
	$(PY) -m src.delta.fetch

extract:          ## data/raw/*.pdf -> data/processed/*.csv + data/sources.csv
	$(PY) -m src.delta.extract

charts:           ## data/processed -> report/figures/*.png
	$(PY) -m src.delta.charts

model:            ## data/processed -> model/delta_model.xlsx (live formulas), recalculated by LibreOffice
	$(PY) -m src.delta.build_excel
	$(PY) -m src.delta.recalc model/delta_model.xlsx

preview:          ## render model sheets to report/figures/model_*.png (README images)
	$(PY) -m src.delta.preview

test:
	$(PY) -m pytest -q tests

clean:
	rm -f data/processed/*.csv
