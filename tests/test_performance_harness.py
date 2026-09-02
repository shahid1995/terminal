import time

import pytest

from benchmarks.performance_harness import BenchmarkRecorder, percentile, summarize_samples


def test_percentile_returns_expected_linear_percentile():
    assert percentile([1, 2, 3, 4, 5], 50) == pytest.approx(3.0)
    assert percentile([1, 2, 3, 4, 5], 95) == pytest.approx(4.8)


def test_percentile_rejects_empty_samples():
    with pytest.raises(ValueError):
        percentile([], 50)


def test_summarize_samples_reports_count_and_percentiles():
    summary = summarize_samples([1, 2, 3, 4, 5])
    assert summary["count"] == 5
    assert summary["min_ms"] == pytest.approx(1.0)
    assert summary["p50_ms"] == pytest.approx(3.0)
    assert summary["p95_ms"] == pytest.approx(4.8)
    assert summary["max_ms"] == pytest.approx(5.0)


def test_benchmark_recorder_records_elapsed_stage_time():
    recorder = BenchmarkRecorder()
    with recorder.measure("stage"):
        time.sleep(0.001)

    report = recorder.report()
    assert report["stage"]["count"] == 1
    assert report["stage"]["min_ms"] >= 0.5


def test_benchmark_recorder_can_record_multiple_samples():
    recorder = BenchmarkRecorder()
    recorder.record("stage", 1.0)
    recorder.record("stage", 3.0)

    report = recorder.report()
    assert report["stage"]["count"] == 2
    assert report["stage"]["p50_ms"] == pytest.approx(2.0)
