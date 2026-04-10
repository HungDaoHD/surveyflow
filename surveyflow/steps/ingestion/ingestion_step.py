"""Ingestion step: definition + rows → rawdata.csv + metadata.json."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from surveyflow.core.base import Step
from surveyflow.steps.ingestion.data_parser import (
    build_encoding_map,
    encode_records,
    enrich_metadata_values,
    parse_rows,
    records_to_dataframe,
)
from surveyflow.steps.ingestion.metadata_parser import CODEABLE_TYPES, parse_metadata
from surveyflow.utils.io import save_csv, save_json

logger = logging.getLogger(__name__)


class IngestionStep(Step):
    """
    Flow
    ----
    1. parse_metadata(definition)              → metadata; values from definition.choices
                                                 if present, else empty {}
    2. parse_rows(pages, definition)           → raw records
    3a. [old API] build_encoding_map + enrich  → derive codes from first-appearance text
    3b. [new API] skip (values already set)
    4. encode_records(records, metadata, map)  → numeric-coded records
    5. records_to_dataframe(records)           → DataFrame
    6. save rawdata.csv + metadata.json

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

        # 1. metadata (values populated from definition.choices)
        logger.info("Parsing metadata …")
        metadata = parse_metadata(definition)

        # 2. raw records (API now returns numeric codes directly)
        profile_status = context.get("profile_status", ["approved"])
        logger.info(
            "Parsing %d row page(s) [filter: %s] …",
            len(rows_pages),
            profile_status or "all",
        )
        raw_records = parse_rows(rows_pages, definition, profile_status)
        logger.info("  → %d respondent rows", len(raw_records))

        # 3. encode: old API → build map from text answers; new API → cast int
        needs_map = any(
            not q["values"]
            for q in metadata["questions"].values()
            if q["answer_type"] in CODEABLE_TYPES
        )
        if needs_map:
            logger.info("Old API format detected — building encoding map from row data …")
            encoding_map = build_encoding_map(raw_records, metadata)
            enrich_metadata_values(metadata, encoding_map)
        else:
            encoding_map = None

        encoded_records = encode_records(raw_records, metadata, encoding_map)

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
