"""xl3 conformance runner CLI — `python -m xl3.runner`.

Implements `conformance/runner-protocol.md`:
  - Iterates fixtures under --fixture-dir
  - Dispatches by fixture kind (static / error / dynamic)
  - Stage 1 cell-value comparison (Stage 2 OOXML comparison NOT implemented)
  - JSON or text report, --filter by tag, --spec-version gate
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .. import __version__, convert, is_xtl_error
from ..types import ConvertOptions, OutputFile
from .compare import compare_workbook_dir, compare_workbooks
from .discover import Fixture, discover_fixtures
from .dynamic import check_utc_today

DEFAULT_FIXTURE_DIR = Path(__file__).resolve().parents[4] / "xl3" / "conformance" / "fixtures"


def _build_argparser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="xl3-py-runner",
        description="xl3 Python — XTL 0.1 conformance runner",
    )
    p.add_argument(
        "--fixture-dir",
        type=Path,
        default=DEFAULT_FIXTURE_DIR,
        help=f"path to conformance/fixtures (default: {DEFAULT_FIXTURE_DIR})",
    )
    p.add_argument(
        "--filter",
        action="append",
        default=[],
        help="only run fixtures with this tag (repeatable)",
    )
    p.add_argument(
        "--id-prefix",
        action="append",
        default=[],
        help="only run fixtures whose id starts with this prefix (repeatable)",
    )
    p.add_argument(
        "--spec-version",
        default="0.1",
        help="run only fixtures whose spec_version <= this value (default: 0.1)",
    )
    p.add_argument(
        "--comparison-stage",
        type=int,
        choices=[1, 2],
        default=1,
        help="active comparison stage; fixtures declaring a higher stage are skipped",
    )
    p.add_argument(
        "--report",
        choices=["json", "text"],
        default="text",
        help="report format",
    )
    return p


def _filter_fixture(fx: Fixture, args: argparse.Namespace) -> str | None:
    """Return a skip reason string, or None to keep."""
    if args.filter:
        if not any(tag in fx.tags for tag in args.filter):
            return f"no tag in {args.filter}"
    if args.id_prefix:
        if not any(fx.id.startswith(p) for p in args.id_prefix):
            return f"id-prefix mismatch ({args.id_prefix})"
    if fx.skip_reason:
        return fx.skip_reason
    if fx.kind == "static" and fx.comparison_stage > args.comparison_stage:
        return f"requires comparison_stage {fx.comparison_stage}"
    if fx.spec_version > args.spec_version:
        return f"requires spec_version {fx.spec_version}"
    return None


def _build_options(fx: Fixture) -> ConvertOptions | None:
    if not fx.inputs:
        return None
    return ConvertOptions(inputs={i.name: i.value for i in fx.inputs})


def _run_static(fx: Fixture, output_files: list[OutputFile]) -> tuple[bool, str]:
    actual_files = [(f.filename, f.data) for f in output_files]
    if fx.expected_path is not None:
        if len(actual_files) != 1:
            return False, f"expected single output file, got {len(actual_files)}"
        return compare_workbooks(actual_files[0][1], fx.expected_path)
    if fx.expected_dir is not None:
        return compare_workbook_dir(actual_files, fx.expected_dir)
    return False, "no expected reference"  # discover.py guards this


def _run_error(fx: Fixture, exc: Exception) -> tuple[bool, str]:
    msg = str(exc)
    if fx.expected_error and fx.expected_error not in msg:
        return False, f"expected error containing {fx.expected_error!r}, got {msg!r}"
    if fx.expected_error_code:
        code = getattr(exc, "code", None)
        if code != fx.expected_error_code:
            return False, f"expected error code {fx.expected_error_code!r}, got {code!r}"
    return True, ""


def run(args: argparse.Namespace) -> int:
    runner_start = datetime.now(timezone.utc)
    fixtures = discover_fixtures(args.fixture_dir)

    results: list[dict[str, Any]] = []
    summary = {"total": 0, "passed": 0, "failed": 0, "skipped": 0}

    for fx in fixtures:
        skip = _filter_fixture(fx, args)
        if skip:
            results.append({"fixture": fx.id, "status": "skip", "reason": skip})
            summary["skipped"] += 1
            continue

        summary["total"] += 1
        if fx.template_path is None or fx.data_path is None:
            results.append(
                {"fixture": fx.id, "status": "fail", "diff": "missing template.xlsx or data.xlsx"}
            )
            summary["failed"] += 1
            continue

        template_bytes = fx.template_path.read_bytes()
        data_bytes = fx.data_path.read_bytes()
        options = _build_options(fx)

        t0 = time.perf_counter()
        try:
            output = convert(template_bytes, data_bytes, options)
        except Exception as exc:  # noqa: BLE001
            duration_ms = int((time.perf_counter() - t0) * 1000)
            if fx.kind == "error":
                passed, diff = _run_error(fx, exc)
            else:
                passed = False
                diff = f"unexpected exception: {exc!r}"
                if not is_xtl_error(exc) and not isinstance(exc, NotImplementedError):
                    # promote unknown exceptions in the diff for visibility
                    diff = f"non-xl3 exception: {type(exc).__name__}: {exc}"
            entry: dict[str, Any] = {
                "fixture": fx.id,
                "status": "pass" if passed else "fail",
                "duration_ms": duration_ms,
            }
            if not passed:
                entry["diff"] = diff
            results.append(entry)
            summary["passed" if passed else "failed"] += 1
            continue

        duration_ms = int((time.perf_counter() - t0) * 1000)
        if fx.kind == "error":
            results.append(
                {
                    "fixture": fx.id,
                    "status": "fail",
                    "duration_ms": duration_ms,
                    "diff": f"expected error {fx.expected_error!r}, got success",
                }
            )
            summary["failed"] += 1
            continue
        if fx.kind == "dynamic":
            passed, diff = check_utc_today(
                [(f.filename, f.data) for f in output],
                fx.dynamic_cells,
                runner_start,
            )
        else:
            passed, diff = _run_static(fx, output)
        entry = {
            "fixture": fx.id,
            "status": "pass" if passed else "fail",
            "duration_ms": duration_ms,
        }
        if not passed:
            entry["diff"] = diff
        results.append(entry)
        summary["passed" if passed else "failed"] += 1

    report = {
        "implementation": "xl3-py",
        "version": __version__,
        "spec_version": args.spec_version,
        "comparison_stage": args.comparison_stage,
        "results": results,
        "summary": summary,
    }
    if args.report == "json":
        json.dump(report, sys.stdout, indent=2)
        sys.stdout.write("\n")
    else:
        _print_text_report(report)
    return 0 if summary["failed"] == 0 else 1


def _print_text_report(report: dict[str, Any]) -> None:
    print(
        f"xl3-py {report['version']} — XTL {report['spec_version']} "
        f"(stage {report['comparison_stage']})"
    )
    print()
    fail_count = 0
    skip_count = 0
    for r in report["results"]:
        status = r["status"]
        line = f"  {status:5}  {r['fixture']}"
        if status == "fail":
            fail_count += 1
            line += f"  -- {r.get('diff', '')}"
        elif status == "skip":
            skip_count += 1
            line += f"  ({r.get('reason', '')})"
        print(line)
    s = report["summary"]
    print()
    print(
        f"summary: {s['passed']}/{s['total']} passed, {s['failed']} failed, "
        f"{s['skipped']} skipped"
    )


def main(argv: list[str] | None = None) -> int:
    args = _build_argparser().parse_args(argv)
    return run(args)


if __name__ == "__main__":
    sys.exit(main())
