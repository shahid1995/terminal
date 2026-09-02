# Terminal Performance Benchmark Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or `superpowers:executing-plans` to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Establish a repeatable, non-invasive baseline for the NIFTY Excel terminal before changing its live data-processing architecture.

**Architecture:** Keep the existing Kite/WebSocket/Excel script untouched in this phase. Add an isolated benchmark package that records stage latency and, on Windows, measures xlwings write and Microsoft Excel calculation latency against the workbook. Results are emitted as JSON so the eventual V2 can be compared against the same metrics.

**Tech Stack:** Python 3, pytest, xlwings on Windows, Microsoft Excel COM, standard-library timing/statistics, optional psutil for RSS memory.

**Spec:** Initial performance-baseline design agreed in chat on 2026-09-02.

## Global Constraints

- Do not modify the existing live Kite/Excel production script in the baseline phase.
- Do not alter StrikeNova or any unrelated repository.
- Do not require live broker credentials for the benchmark helpers.
- Microsoft Excel COM measurements are Windows-only; do not substitute LibreOffice timings as authoritative Excel results.
- Benchmark artifacts must not overwrite the user's original workbook unintentionally.

---

### Task 1: Benchmark metric primitives

**Files:**
- Create: `benchmarks/performance_harness.py`
- Create: `benchmarks/__init__.py`
- Test: `tests/test_performance_harness.py`

**Interfaces:**
- `percentile(samples, p) -> float`
- `summarize_samples(samples) -> dict`
- `BenchmarkRecorder.record(stage, elapsed_ms)`
- `BenchmarkRecorder.measure(stage)` context manager
- `BenchmarkRecorder.report() -> dict`

- [x] Write the failing tests.
- [x] Verify the tests fail before implementation.
- [x] Implement the minimal metric primitives.
- [x] Verify the metric tests pass.

### Task 2: Windows Excel benchmark entry point

**Files:**
- Modify: `benchmarks/performance_harness.py`
- Test: `tests/test_performance_harness.py`

**Interfaces:**
- CLI: `python -m benchmarks.performance_harness --workbook <path> [--iterations N] [--rows N] [--columns N] [--output <path>]`
- JSON output containing workbook, iteration shape, memory, and stage percentile summaries.

- [x] Implement an isolated temporary workbook copy so the user's original is never opened or modified by the benchmark.
- [x] Target the real `TickData!A:AL` write path, including the header row produced by `xlwings(...).value` with `index=False`.
- [x] Measure Excel write latency and full workbook calculation latency separately.
- [x] Record P50/P95/P99 and memory observations.
- [x] Disable workbook macros and external-link updates during the benchmark.
- [ ] Run on a Windows machine with Microsoft Excel and the uploaded workbook copy.

### Task 3: Baseline comparison

**Files:**
- Create: `docs/benchmarks/` result artifact only after execution

- [ ] Run the benchmark against the current workbook before architecture changes.
- [ ] Preserve the raw JSON results.
- [ ] Use the baseline to set acceptance targets for the V2 implementation.

### Task 4: Verification gate

- [x] Confirm the production script has no source changes in the benchmark branch.
- [ ] Run all available tests on the benchmark branch.
- [ ] Review benchmark output for plausible sample counts and outliers.
- [ ] Do not claim performance improvement until V1 and V2 are measured under equivalent conditions.
