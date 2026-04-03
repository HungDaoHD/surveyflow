"""Pipeline configuration."""

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class PipelineConfig:
    """Configuration for a SurveyFlow pipeline run.

    Parameters
    ----------
    survey_id : int
        Qme survey ID to process.
    output_dir : str
        Directory where rawdata.csv, metadata.json (and later datatable.xlsx)
        will be saved. Created automatically if it does not exist.
    profile_status : list[str]
        Filter rows by task profile status.
        Allowed values: "pending", "approved", "rejected", "screen_out".
        Empty list means no filter (all statuses included).
    datatable_config : str | None
        Path to datatable.json config file.  Required for the table step.
    """

    survey_id: int
    output_dir: str = "."
    profile_status: list[str] = field(default_factory=list)
    datatable_config: Optional[str] = None
