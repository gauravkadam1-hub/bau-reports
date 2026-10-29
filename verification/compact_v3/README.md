# Current build 2026-10-24.3 evidence

111 regression tests passed; 24 full Polars/FastExcel tests are blocked by missing dependencies. Desktop Excel was not available.

The size JSONs record identical synthetic datasets through the original engine, working v2 writer, v3 writer before additional compression (v3-uncompressed), and complete v3 writer. Timing scopes differ as documented; do not compare original end-to-end timing with writer-only timing. Large XLSX outputs are omitted to keep the project package small.

The benchmark_worker.py accepts --workspace PATH --mode original|v2|v3|v3-uncompressed --shape sparse|dense --rows 20000. Workspace must contain original/bau-reports-main, v2/bau-reports-main and updated/bau-reports-main; mkdir workspace/verification first. These are synthetic data, and the workbook-writer tests deliberately use a Python reference report, not a simulated Polars backend.

parity JSONs compare populated values/formulas and cached results. recalculation.json records an independent recalculation of 224 formula caches from the separate synthetic sample. integrity.json verifies unchanged original inputs and other engines. The strict runner's BLOCKED result is expected until the real dataframe/Excel-reader dependencies are installed.
