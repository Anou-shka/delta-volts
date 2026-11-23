# Rebuild everything from primary sources. `make all` runs every stage in order.
PY ?= python3

.PHONY: all data fetch extract market charts model preview report test clean

all: data model charts report

data: fetch extract

fetch:            ## download filings into data/raw/ (skips files already present)
	$(PY) -m src.delta.fetch

extract:          ## data/raw/*.pdf -> data/processed/*.csv + data/sources.csv
	$(PY) -m src.delta.extract

charts:           ## data/processed -> report/figures/*.png
	$(PY) -m src.delta.charts

market:           ## refresh the market-data snapshot (prices, peers, FX) - changes the valuation date
	$(PY) -m src.delta.market

model:            ## Python mirror (valuation_summary.json) -> Excel model with live formulas, recalculated
	$(PY) -m src.delta.model
	$(PY) -m src.delta.build_excel
	$(PY) -m src.delta.recalc model/delta_model.xlsx

preview:          ## render model sheets to report/figures/model_*.png (README images)
	$(PY) -m src.delta.preview

report:           ## model outputs -> report/delta_initiation.md (every number generated from the model)
	$(PY) -m src.delta.report

test:
	$(PY) -m pytest -q tests

clean:
	rm -f data/processed/*.csv
