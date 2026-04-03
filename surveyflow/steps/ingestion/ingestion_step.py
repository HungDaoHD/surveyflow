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
from surveyflow.steps.ingestion.metadata_parser import (
    build_encoding_map,
    enrich_metadata_values,
    parse_metadata,
)
from surveyflow.utils.io import save_csv, save_json

logger = logging.getLogger(__name__)


class IngestionStep(Step):
    """
    Flow
    ----
    1. parse_metadata(definition)              → metadata skeleton (values = {})
    2. parse_rows(pages, definition)           → raw-text records
    3. build_encoding_map(records, metadata)   → {label: {text → code}}
    4. enrich_metadata_values(metadata, map)   → values = {code: text}
    5. encode_records(records, map, metadata)  → numeric-coded records
    6. records_to_dataframe(records)           → DataFrame
    7. save rawdata.csv + metadata.json

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

        # 1. metadata skeleton
        logger.info("Parsing metadata …")
        metadata = parse_metadata(definition)

        # 2. raw-text records  (default: approved only)
        profile_status = context.get("profile_status", ["approved"])
        logger.info(
            "Parsing %d row page(s) [filter: %s] …",
            len(rows_pages),
            profile_status or "all",
        )
        raw_records = parse_rows(rows_pages, definition, profile_status)
        logger.info("  → %d respondent rows", len(raw_records))

        # 3. build encoding map  (text → 1-based codes, derived from row data)
        encoding_map = build_encoding_map(raw_records, metadata)

        # 4. enrich metadata  values = {code_str: label_text}
        enrich_metadata_values(metadata, encoding_map)

        # 5. encode records  (SA → int, MA/ranking → "1;3;5")
        encoded_records = encode_records(raw_records, encoding_map, metadata)

        # 6. DataFrame
        df = records_to_dataframe(encoded_records, definition, metadata)

        # 7. save
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
