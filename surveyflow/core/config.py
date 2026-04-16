"""Pipeline configuration."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class PipelineConfig:
    """Configuration for a SurveyFlow pipeline run.

    Parameters
    ----------
    definition : dict
        Raw response from get_survey_definition (question structure).
    rows_pages : list[dict]
        Paginated raw responses from get_survey_rows (one dict per page).
    output_dir : str
        Directory where rawdata.csv, metadata.json and datatable.xlsx are saved.
        Created automatically if it does not exist.
    profile_status : list[str]
        Filter rows by task profile status.
        Allowed: "pending", "approved", "rejected", "screen_out".
        Default ["approved"]. Empty list = no filter (all statuses).
    datatable_config : str | dict | None
        Path to datatable.json  OR  already-parsed dict.
        When provided the TableStep runs and produces datatable.xlsx.
        When None only rawdata.csv + metadata.json are produced.
    """

    definition:       dict
    rows_pages:       list[dict]
    output_dir:       str       = "."
    profile_status:    list[str]  = field(default_factory=lambda: ["approved"])
    datatable_config:  Optional[str | dict] = None
    version:           Optional[str] = None
    """Version tag for the output folder.
    - If provided (e.g. "v1", "v2"): output goes to ``{output_dir}/v1/``
    - If None: auto-generates a timestamp folder ``{output_dir}/20260408_091623/``
    """
