"""Performance benchmark helpers for the NIFTY Excel terminal.

This module is intentionally independent from the live Kite/Excel script. It
provides reusable timing/reporting primitives plus a Windows-only Excel
benchmark that measures workbook write and recalculation latency without
changing workbook formulas or the live market-data pipeline.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import platform
import statistics
import time
from collections import defaultdict
from contextlib import contextmanager
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Sequence


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


def _excel_benchmark(workbook_path: Path, iterations: int, rows: int, columns: int) -> Dict[str, object]:
    """Measure xlwings write and Excel calculation latency on Windows.

    The benchmark uses a temporary worksheet in the target workbook and
    deletes it before returning. It never touches the existing terminal sheets.
    """
    if platform.system() != "Windows":
        raise RuntimeError("Excel COM benchmark requires Windows with Microsoft Excel installed")

    import xlwings as xw

    workbook = xw.Book(str(workbook_path))
    recorder = BenchmarkRecorder()
    sheet_name = "__PERF_BENCHMARK__"
    sheet = workbook.sheets.add(sheet_name, after=workbook.sheets[-1])
    try:
        values = [[(row * columns) + column for column in range(columns)] for row in range(rows)]
        target = sheet.range("A1").resize(rows, columns)

        for _ in range(iterations):
            with recorder.measure("excel_write_ms"):
                target.value = values
            with recorder.measure("excel_calculation_ms"):
                workbook.app.calculate()

        return {
            "workbook": str(workbook_path),
            "iterations": iterations,
            "rows": rows,
            "columns": columns,
            "memory_mb": _memory_mb(),
            "stages": recorder.report(),
        }
    finally:
        sheet.delete()
        workbook.save()
        workbook.close()


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Benchmark Excel terminal performance")
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--iterations", type=int, default=20)
    parser.add_argument("--rows", type=int, default=300)
    parser.add_argument("--columns", type=int, default=38)
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
