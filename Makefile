# Rebuild everything from primary sources. `make all` runs every stage in order.
PY ?= python3

.PHONY: all data fetch extract charts test clean

all: data charts

data: fetch extract

fetch:            ## download filings into data/raw/ (skips files already present)
	$(PY) -m src.delta.fetch

extract:          ## data/raw/*.pdf -> data/processed/*.csv + data/sources.csv
	$(PY) -m src.delta.extract

charts:           ## data/processed -> report/figures/*.png
	$(PY) -m src.delta.charts

test:
	$(PY) -m pytest -q tests

clean:
	rm -f data/processed/*.csv
