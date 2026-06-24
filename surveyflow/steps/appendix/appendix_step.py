"""AppendixStep: chart_data.json → slides.pptx."""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from surveyflow.core.base import Step

logger = logging.getLogger(__name__)


class AppendixStep(Step):
    """Generate a PowerPoint chart appendix from chart_data.json.

    Reads
    -----
    context["chart_data_path"] : str
        Path to chart_data.json written by TableStep.render_chart_data().
    context["output_dir"] : str
        Versioned output directory (e.g. output/VN8947/v1/).
        slides.pptx is written here as a sibling of datatable.xlsx.

    Writes
    ------
    context["slides_path"] : str
        Absolute path of the generated slides.pptx.
    """

    def run(self, context: dict[str, Any]) -> dict[str, Any]:
        chart_data_path = context.get("chart_data_path")
        if not chart_data_path:
            raise RuntimeError(
                "AppendixStep requires chart_data_path in context. "
                "Run TableStep.render_chart_data() first."
            )
        if not Path(chart_data_path).exists():
            raise FileNotFoundError(f"chart_data.json not found: {chart_data_path}")

        slides_path = str(Path(chart_data_path).parent / "slides.pptx")

        logger.info("AppendixStep: generating slides -> %s", slides_path)
        try:
            from surveyflow.generate_pptx import generate
        except ImportError:
            raise ImportError(
                "python-pptx is required for AppendixStep. "
                "Install it with: pip install python-pptx"
            )

        generate(chart_data_path, slides_path)
        context["slides_path"] = slides_path
        logger.info("slides.pptx -> %s", slides_path)
        return context
