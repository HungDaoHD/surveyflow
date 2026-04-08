"""SurveyFlow Pipeline — orchestrates Ingestion → Table."""
from __future__ import annotations

import json
import logging
from datetime import datetime
from pathlib import Path

from surveyflow.core.config import PipelineConfig
from surveyflow.steps.ingestion.ingestion_step import IngestionStep
from surveyflow.steps.table.table_step import TableStep

logger = logging.getLogger(__name__)


class Pipeline:
    """Run the full survey data pipeline.

    Usage
    -----
    >>> pipeline = Pipeline(PipelineConfig(
    ...     definition  = definition,   # from get_survey_definition
    ...     rows_pages  = rows_pages,   # from get_survey_rows (all pages)
    ...     output_dir  = "output/VN8947",
    ...     datatable_config = "output/VN8947/datatable.json",   # optional
    ... ))
    >>> result = pipeline.run()
    >>> result["rawdata_path"]      # "output/VN8947/rawdata.csv"
    >>> result["metadata_path"]     # "output/VN8947/metadata.json"
    >>> result["datatable_path"]    # "output/VN8947/datatable.xlsx" (if config given)
    """

    def __init__(self, config: PipelineConfig) -> None:
        self.config = config

    # ── public ────────────────────────────────────────────────────────────────

    def run(self) -> dict:
        """Execute all steps and return the final context dict."""
        cfg = self.config

        version    = cfg.version or datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = str(Path(cfg.output_dir) / version)

        logger.info("Pipeline started  →  %s", output_dir)

        context: dict = {
            "definition":     cfg.definition,
            "rows_pages":     cfg.rows_pages,
            "output_dir":     output_dir,
            "profile_status": cfg.profile_status,
            "version":        version,
        }

        # ── Step 1: Ingestion ──────────────────────────────────────────
        logger.info("── Step 1: Ingestion")
        context = IngestionStep().run(context)

        # ── Step 2: Table (optional) ───────────────────────────────────
        if cfg.datatable_config is not None:
            logger.info("── Step 2: Table")
            context["df"]              = context["rawdata"]   # bridge key
            context["datatable_config"] = self._load_datatable_config(
                cfg.datatable_config
            )
            context = TableStep().run(context)

        logger.info("Pipeline complete.")
        return context

    # ── internal ──────────────────────────────────────────────────────────────

    @staticmethod
    def _load_datatable_config(source: str | dict) -> dict:
        if isinstance(source, dict):
            return source
        path = Path(source)
        if not path.exists():
            raise FileNotFoundError(f"datatable_config not found: {path}")
        with path.open(encoding="utf-8") as f:
            return json.load(f)
