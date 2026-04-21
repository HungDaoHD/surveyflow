"""Ingestion step: definition + rows → rawdata.csv + metadata.json."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from surveyflow.core.base import Step
from surveyflow.steps.ingestion.data_parser import (
    encode_records,
    parse_rows,
    records_to_dataframe,
)
from surveyflow.steps.ingestion.metadata_parser import parse_metadata
from surveyflow.utils.io import save_csv, save_json

logger = logging.getLogger(__name__)


class IngestionStep(Step):
    """
    Flow
    ----
    1. parse_metadata(definition)        → metadata keyed by str(question_id)
    2. parse_rows(pages, definition)     → raw text/code records
    3. encode_records(records, metadata) → SA → int, NUM → number, rest as-is
    4. records_to_dataframe(records)     → ordered DataFrame
    5. save rawdata.csv + metadata.json

    Context inputs
    --------------
    definition   : dict   raw get_survey_definition response
    rows_pages   : list   raw get_survey_rows responses (one per page)
    output_dir   : str    destination folder

    Context outputs
    ---------------
    rawdata        : pd.DataFrame
    metadata       : dict
    rawdata_path   : str
    metadata_path  : str
    """

    def run(self, context: dict[str, Any]) -> dict[str, Any]:
        definition = context["definition"]
        rows_pages = context["rows_pages"]
        output_dir = Path(context.get("output_dir", "."))
        output_dir.mkdir(parents=True, exist_ok=True)

        # 1. metadata
        logger.info("Parsing metadata …")
        metadata = parse_metadata(definition)

        # 2. raw records (API returns numeric codes directly)
        profile_status = context.get("profile_status", ["approved"])
        logger.info(
            "Parsing %d row page(s) [filter: %s] …",
            len(rows_pages),
            profile_status or "all",
        )
        # Build {question_id: [other_choice_codes]} for other_text column naming
        other_codes_map: dict[int, list[str]] = {
            meta["question_id"]: meta["other_choice_codes"]
            for meta in metadata["questions"].values()
            if meta.get("other_choice_codes")
        }
        raw_records = parse_rows(
            rows_pages, definition, profile_status,
            other_codes_map or None,
        )
        logger.info("  → %d respondent rows", len(raw_records))

        # 3. encode answers to numeric codes
        encoded_records = encode_records(raw_records, metadata)

        # 4. DataFrame
        df = records_to_dataframe(encoded_records, definition, metadata)

        # 6. save
        rawdata_path  = output_dir / "rawdata.csv"
        metadata_path = output_dir / "metadata.json"
        save_csv(df, rawdata_path)
        save_json(metadata, metadata_path)

        logger.info("Saved → %s", rawdata_path)
        logger.info("Saved → %s", metadata_path)

        context["rawdata"]       = df
        context["metadata"]      = metadata
        context["rawdata_path"]  = str(rawdata_path)
        context["metadata_path"] = str(metadata_path)
        return context
        