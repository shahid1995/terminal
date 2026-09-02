"""Performance benchmark helpers for the NIFTY Excel terminal.

This module is intentionally independent from the live Kite/Excel script. It
provides reusable timing/reporting primitives plus a Windows-only Excel
benchmark that measures the real TickData write and workbook recalculation
path on an isolated copy of the supplied workbook.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import platform
import shutil
import statistics
import tempfile
import time
from collections import defaultdict
from contextlib import contextmanager
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Sequence


TICKDATA_COLUMNS = 38  # Current Python terminal writes A:AL.
TICKDATA_SHEET = "TickData"


def percentile(samples: Sequence[float], p: float) -> float:
    """Return a linearly interpolated percentile from numeric samples."""
    if not samples:
        raise ValueError("samples must not be empty")
    if not 0 <= p <= 100:
        raise ValueError("p must be between 0 and 100")

    ordered = sorted(float(value) for value in samples)
    if len(ordered) == 1:
        return ordered[0]

    rank = (len(ordered) - 1) * (p / 100.0)
    lower = math.floor(rank)
    upper = math.ceil(rank)
    if lower == upper:
        return ordered[lower]
    weight = rank - lower
    return ordered[lower] + (ordered[upper] - ordered[lower]) * weight


def summarize_samples(samples: Sequence[float]) -> Dict[str, float]:
    """Summarize elapsed-time samples in milliseconds."""
    if not samples:
        raise ValueError("samples must not be empty")
    values = [float(value) for value in samples]
    return {
        "count": len(values),
        "min_ms": min(values),
        "p50_ms": percentile(values, 50),
        "p95_ms": percentile(values, 95),
        "p99_ms": percentile(values, 99),
        "mean_ms": statistics.fmean(values),
        "max_ms": max(values),
    }


class BenchmarkRecorder:
    """Collect named elapsed-time samples and emit percentile summaries."""

    def __init__(self) -> None:
        self._samples: Dict[str, List[float]] = defaultdict(list)

    def record(self, stage: str, elapsed_ms: float) -> None:
        if elapsed_ms < 0:
            raise ValueError("elapsed_ms must not be negative")
        self._samples[stage].append(float(elapsed_ms))

    @contextmanager
    def measure(self, stage: str) -> Iterator[None]:
        started = time.perf_counter()
        try:
            yield
        finally:
            self.record(stage, (time.perf_counter() - started) * 1000.0)

    def report(self) -> Dict[str, Dict[str, float]]:
        return {stage: summarize_samples(samples) for stage, samples in self._samples.items()}


def _memory_mb() -> float | None:
    """Return current RSS in MB when psutil is installed."""
    try:
        import psutil

        return psutil.Process(os.getpid()).memory_info().rss / (1024 * 1024)
    except ImportError:
        return None


def _prepare_workbook_copy(source: Path, directory: Path) -> Path:
    """Copy a workbook into an isolated benchmark directory."""
    source = source.resolve()
    directory.mkdir(parents=True, exist_ok=True)
    destination = directory / source.name
    shutil.copy2(source, destination)
    return destination


def _build_tickdata_matrix(data_rows: int, columns: int) -> List[List[object]]:
    """Build the same header-plus-data shape written by xlwings ``index=False``."""
    header = [f"column_{column + 1}" for column in range(columns)]
    data = [
        [
            (row + 1) * (column + 1)
            for column in range(columns)
        ]
        for row in range(data_rows)
    ]
    return [header, *data]


def _excel_benchmark(workbook_path: Path, iterations: int, rows: int, columns: int) -> Dict[str, object]:
    """Measure the real TickData write and Excel recalculation path on Windows.

    The supplied workbook is copied to a temporary directory first. The copy's
    ``TickData!A1:AL{rows + 1}`` area is overwritten with representative header
    and data values, Excel recalculation is triggered, and the temporary
    workbook is discarded. The original workbook is never opened or modified.
    """
    if platform.system() != "Windows":
        raise RuntimeError("Excel COM benchmark requires Windows with Microsoft Excel installed")
    if columns != TICKDATA_COLUMNS:
        raise ValueError(f"columns must be {TICKDATA_COLUMNS} for the current TickData writer")

    import xlwings as xw

    recorder = BenchmarkRecorder()
    values = _build_tickdata_matrix(rows, columns)

    with tempfile.TemporaryDirectory(prefix="terminal-perf-") as temp_dir:
        benchmark_copy = _prepare_workbook_copy(workbook_path, Path(temp_dir))
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        app.api.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        workbook = None
        try:
            workbook = app.books.open(str(benchmark_copy), update_links=False, read_only=False)
            sheet = workbook.sheets[TICKDATA_SHEET]
            target = sheet.range("A1").resize(rows + 1, columns)

            for _ in range(iterations):
                with recorder.measure("excel_write_ms"):
                    target.value = values
                with recorder.measure("excel_calculation_ms"):
                    app.calculate()

            return {
                "workbook": str(workbook_path),
                "benchmark_copy": str(benchmark_copy),
                "sheet": TICKDATA_SHEET,
                "write_columns": columns,
                "data_rows": rows,
                "write_rows_including_header": rows + 1,
                "iterations": iterations,
                "memory_mb": _memory_mb(),
                "stages": recorder.report(),
            }
        finally:
            if workbook is not None:
                workbook.close(save=False)
            app.quit()


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Benchmark Excel terminal performance")
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--iterations", type=int, default=20)
    parser.add_argument("--rows", type=int, default=300)
    parser.add_argument("--columns", type=int, default=TICKDATA_COLUMNS)
    parser.add_argument("--output", type=Path, default=Path("benchmark-results.json"))
    args = parser.parse_args(list(argv) if argv is not None else None)

    if not args.workbook.exists():
        parser.error(f"workbook does not exist: {args.workbook}")
    if args.iterations < 1 or args.rows < 1 or args.columns < 1:
        parser.error("iterations, rows and columns must all be positive")

    result = _excel_benchmark(args.workbook, args.iterations, args.rows, args.columns)
    args.output.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
