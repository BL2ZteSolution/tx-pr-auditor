"""Repeatable synthetic capacity benchmark for the manifest workload limit."""
from __future__ import annotations

import argparse
import json
import time
import tracemalloc

from tests.test_audit_final_po import expected_record, final_record, run_pipeline_for_records


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, default=10000)
    args = parser.parse_args()
    finals = [
        final_record(
            source_row=index + 2,
            site_code=f"SITE-{index:06d}",
            du=f"DU-{index:06d}",
            request_number=f"REQ-{index:06d}",
            dispatch_order_number=f"DO-{index:06d}",
        )
        for index in range(args.rows)
    ]
    expected = [
        expected_record(
            source_row=index + 2,
            site_code=f"SITE-{index:06d}",
            du=f"DU-{index:06d}",
        )
        for index in range(args.rows)
    ]
    tracemalloc.start()
    started = time.perf_counter()
    results = run_pipeline_for_records(finals, expected)
    elapsed = time.perf_counter() - started
    _current, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    payload = {
        "rows": args.rows,
        "resultRows": len(results),
        "elapsedSeconds": round(elapsed, 3),
        "peakTracedMiB": round(peak / 1024 / 1024, 2),
        "cancellationProbeIntervalRows": 250,
        "progressHeartbeatSeconds": 30,
    }
    print(json.dumps(payload, indent=2))
    if len(results) != args.rows:
        raise SystemExit("Benchmark result count mismatch.")


if __name__ == "__main__":
    main()
