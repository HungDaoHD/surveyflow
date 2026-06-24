"""SurveyFlow Pipeline — orchestrates Ingestion → Table."""
from __future__ import annotations

import json
import logging
from datetime import datetime
from pathlib import Path

import pandas as pd

from surveyflow.core.config import PipelineConfig
from surveyflow.steps.ingestion.ingestion_step import IngestionStep
from surveyflow.steps.table.table_step import TableStep

logger = logging.getLogger(__name__)


class Pipeline:
    """Run the full survey data pipeline.

    Folder layout
    -------------
    output_dir/
    ├── data/           ← rawdata.csv + metadata.json (generated once from mcp/)
    ├── datatable/      ← datatable.json (managed by Claude)
    ├── v1/             ← datatable.xlsx only
    └── v2/

    Usage — first run (ingestion + table)
    -------------------------------------
    >>> pipeline = Pipeline(PipelineConfig(
    ...     definition       = definition,
    ...     rows_pages       = rows_pages,
    ...     output_dir       = "output/VN8947",
    ...     datatable_config = "output/VN8947/datatable/datatable.json",
    ...     version          = "v1",
    ... ))
    >>> result = pipeline.run()

    Usage — table-only (data already exists)
    -----------------------------------------
    >>> pipeline = Pipeline(PipelineConfig(
    ...     output_dir       = "output/VN8947",
    ...     skip_ingestion   = True,
    ...     datatable_config = "output/VN8947/datatable/datatable.json",
    ...     version          = "v2",
    ... ))
    >>> result = pipeline.run()
    """

    def __init__(self, config: PipelineConfig) -> None:
        self.config = config

    # ── public ────────────────────────────────────────────────────────────────

    def run(self) -> dict:
        """Execute pipeline steps and return the final context dict."""
        cfg = self.config

        version      = cfg.version or datetime.now().strftime("%Y%m%d_%H%M%S")
        base_dir     = Path(cfg.output_dir)
        data_dir     = Path(cfg.data_dir) if cfg.data_dir else base_dir / "data"
        versioned_dir = base_dir / version   # datatable.xlsx goes here

        logger.info("Pipeline started  →  %s", versioned_dir)

        context: dict = {
            "definition":     cfg.definition,
            "rows_pages":     cfg.rows_pages,
            "export_df":      cfg.export_df,
            "data_dir":       str(data_dir),
            "output_dir":     str(versioned_dir),   # table step writes xlsx here
            "profile_status": cfg.profile_status,
            "skip_render":    cfg.skip_render,
            "table_indices":  cfg.table_indices,
            "version":        version,
            "lang":           cfg.lang,
        }

        # ── Metadata-only mode ────────────────────────────────────────────
        # Parse definition → save metadata.json immediately, skip rawdata.
        # Useful for building datatable artifacts before data_export.csv is ready.
        if cfg.metadata_only:
            if cfg.definition is None:
                raise ValueError("definition is required for metadata_only mode")
            from surveyflow.steps.ingestion.metadata_parser import parse_metadata
            from surveyflow.utils.io import save_json
            import json as _json
            data_dir.mkdir(parents=True, exist_ok=True)
            metadata      = parse_metadata(cfg.definition)
            metadata_path = data_dir / "metadata.json"
            save_json(metadata, metadata_path)
            logger.info("metadata_only: saved → %s", metadata_path)

            # If rawdata.csv already exists (survey re-run), annotate rawdata_columns
            # so the metadata remains fully functional for table/quality steps.
            rawdata_path = data_dir / "rawdata.csv"
            if rawdata_path.exists():
                from surveyflow.steps.ingestion.export_parser import annotate_rawdata_columns
                annotate_rawdata_columns(metadata_path, rawdata_path)
                metadata = _json.loads(metadata_path.read_text(encoding="utf-8"))
                logger.info("metadata_only: annotated rawdata_columns from existing rawdata")

            context["metadata"]      = metadata
            context["metadata_path"] = str(metadata_path)
            return context

        # ── Step 1: Ingestion ──────────────────────────────────────────────
        if not cfg.skip_ingestion:
            if cfg.definition is None:
                raise ValueError(
                    "definition is required when skip_ingestion=False"
                )
            if cfg.rows_pages is None and cfg.export_df is None:
                raise ValueError(
                    "rows_pages or export_df is required when skip_ingestion=False"
                )
            logger.info("── Step 1: Ingestion")
            context = IngestionStep().run(context)
        else:
            # Load existing rawdata + metadata from data_dir
            rawdata_path  = data_dir / "rawdata.csv"
            metadata_path = data_dir / "metadata.json"
            if not rawdata_path.exists():
                raise FileNotFoundError(
                    f"rawdata.csv not found in {data_dir}. "
                    "Run pipeline without skip_ingestion first."
                )
            logger.info("── Step 1: Ingestion (skipped — reading from %s)", data_dir)
            context["rawdata"]       = pd.read_csv(rawdata_path)
            with metadata_path.open(encoding="utf-8") as f:
                context["metadata"]  = json.load(f)
            context["rawdata_path"]  = str(rawdata_path)
            context["metadata_path"] = str(metadata_path)

        # ── Step 2: Quality check (optional, standalone) ──────────────────
        if cfg.run_quality:
            from surveyflow.steps.quality.quality_step import QualityStep
            logger.info("── Step 2: Quality check")
            context["base_output_dir"] = str(base_dir)
            context = QualityStep().run(context)

        # ── Step 3: Table (optional) ───────────────────────────────────────
        if cfg.datatable_config is not None:
            logger.info("── Step 2: Table — compute")
            context["df"]               = context["rawdata"]
            context["datatable_config"] = self._load_datatable_config(
                cfg.datatable_config
            )
            table_step = TableStep()
            context = table_step.compute(context)
            if not cfg.skip_render:
                logger.info("── Step 2: Table — render xlsx")
                context = table_step.render_xlsx(context)
                logger.info("── Step 2: Table — render chart_data")
                context = table_step.render_chart_data(context)

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
