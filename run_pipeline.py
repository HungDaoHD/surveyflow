"""Generic pipeline runner — reads pre-fetched MCP data from JSON files.

Usage (Claude calls this after fetching MCP data):
    python run_pipeline.py \
        --input-dir   output/VN8947/input   \
        --output-dir  output/VN8947         \
        --version     v1                    \
        [--datatable-config datatable.json] \
        [--profile-status approved]
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from pathlib import Path

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
sys.path.insert(0, str(Path(__file__).parent))

from surveyflow import Pipeline, PipelineConfig


def _load_input(input_dir: Path) -> tuple[dict, list[dict]]:
    """Load definition.json + rows_page_*.json from input_dir."""
    def_path = input_dir / "definition.json"
    if not def_path.exists():
        raise FileNotFoundError(f"definition.json not found in {input_dir}")

    with def_path.open(encoding="utf-8") as f:
        definition = json.load(f)

    rows_files = sorted(input_dir.glob("rows_page_*.json"))
    if not rows_files:
        raise FileNotFoundError(f"No rows_page_*.json found in {input_dir}")

    rows_pages = []
    for p in rows_files:
        with p.open(encoding="utf-8") as f:
            rows_pages.append(json.load(f))

    logging.info("Loaded definition + %d row page(s) from %s", len(rows_pages), input_dir)
    return definition, rows_pages


def main() -> None:
    parser = argparse.ArgumentParser(description="SurveyFlow Pipeline Runner")
    parser.add_argument("--input-dir",        required=True,  help="Folder with definition.json + rows_page_*.json")
    parser.add_argument("--output-dir",       required=True,  help="Base output folder (version subfolder created inside)")
    parser.add_argument("--version",          default=None,   help="Version tag e.g. v1 (default: auto timestamp)")
    parser.add_argument("--datatable-config", required=True,  help="Path to datatable.json config")
    parser.add_argument("--profile-status",   default="approved", help="Comma-separated statuses e.g. approved,pending")
    args = parser.parse_args()

    input_dir  = Path(args.input_dir)
    output_dir = Path(args.output_dir)
    statuses   = [s.strip() for s in args.profile_status.split(",") if s.strip()]

    definition, rows_pages = _load_input(input_dir)

    result = Pipeline(PipelineConfig(
        definition       = definition,
        rows_pages       = rows_pages,
        output_dir       = str(output_dir),
        profile_status   = statuses,
        version          = args.version,
        datatable_config = args.datatable_config,
    )).run()

    print("\n--- Output ---")
    print(f"  rawdata  : {result['rawdata_path']}")
    print(f"  metadata : {result['metadata_path']}")
    if "datatable_path" in result:
        print(f"  datatable: {result['datatable_path']}")
    print(f"  rows     : {result['rawdata'].shape[0]}")
    print(f"  version  : {result['version']}")


if __name__ == "__main__":
    main()
