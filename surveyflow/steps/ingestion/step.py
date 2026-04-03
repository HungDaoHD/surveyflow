"""Ingestion step: definition + rows → rawdata.csv + metadata.json."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from surveyflow.core.base import Step
from surveyflow.steps.ingestion.data_parser import parse_rows, records_to_dataframe
from surveyflow.steps.ingestion.metadata_parser import (
    parse_metadata,
    enrich_metadata_values,
)
from surveyflow.utils.io import save_csv, save_json

logger = logging.getLogger(__name__)


class IngestionStep(Step):
    """Converts raw Qme API responses to rawdata.csv and metadata.json.

    Expected context keys (inputs)
    --------------------------------
    ``definition``
        Raw dict from ``get_survey_definition``.
    ``rows_pages``
        List of raw dicts from ``get_survey_rows`` (one per page).

    Added context keys (outputs)
    -----------------------------
    ``rawdata``
        pandas DataFrame — one row per respondent.
    ``metadata``
        Enriched metadata dict.
    ``rawdata_path``
        Absolute path to the saved CSV file.
    ``metadata_path``
        Absolute path to the saved JSON file.
    """

    def run(self, context: dict[str, Any]) -> dict[str, Any]:
        definition  = context["definition"]
        rows_pages  = context["rows_pages"]
        output_dir  = Path(context.get("output_dir", "."))
        output_dir.mkdir(parents=True, exist_ok=True)

        # ── 1. metadata ────────────────────────────────────────────────
        logger.info("Parsing metadata …")
        metadata = parse_metadata(definition)

        # ── 2. rawdata ─────────────────────────────────────────────────
        logger.info("Parsing %d row page(s) …", len(rows_pages))
        records = parse_rows(rows_pages, definition)
        logger.info("  → %d respondent rows", len(records))

        # ── 3. enrich metadata with observed answer values ─────────────
        enrich_metadata_values(metadata, records)

        # ── 4. convert to DataFrame ────────────────────────────────────
        df = records_to_dataframe(records, definition)

        # ── 5. save files ──────────────────────────────────────────────
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
