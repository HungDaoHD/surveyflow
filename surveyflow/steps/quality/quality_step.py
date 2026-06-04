"""Quality check step — validates show_condition and contradiction_settings."""
from __future__ import annotations

import json
import logging
from collections import Counter
from pathlib import Path
from typing import Any

import pandas as pd

from surveyflow.core.base import Step
from surveyflow.steps.quality.checker import (
    build_label_to_q,
    describe_condition,
    describe_trigger,
    eval_condition_vec,
    has_answer_vec,
)

logger = logging.getLogger(__name__)

_PROFILE_ID_COL = "profile_id"


class QualityStep(Step):
    """Validate show_condition and contradiction_settings for every respondent.

    Context inputs
    --------------
    rawdata         : pd.DataFrame
    metadata        : dict
    base_output_dir : str   base survey folder (quality/ is created inside)

    Context outputs
    ---------------
    quality_report      : dict
    quality_dir         : str
    quality_report_path : str
    quality_csv_path    : str
    """

    def run(self, context: dict[str, Any]) -> dict[str, Any]:
        df       = context["rawdata"]
        metadata = context["metadata"]

        # Quality output goes to base_dir/quality/ (not versioned subfolder)
        base_dir    = Path(context.get("base_output_dir") or context.get("output_dir", "."))
        quality_dir = base_dir / "quality"
        quality_dir.mkdir(parents=True, exist_ok=True)

        questions  = metadata.get("questions", {})
        label_to_q = build_label_to_q(questions)

        # Identify profile-ID column
        pid_col = _PROFILE_ID_COL if _PROFILE_ID_COL in df.columns else df.columns[0]

        # Answer types that are trackable in rawdata (have columns to check)
        _TRACKABLE_TYPES = {"SA", "MA", "Matrix_SA", "Matrix_MA", "Matrix_NUM", "ranking"}

        violations: list[dict] = []
        n_sc      = 0   # questions with show_condition
        n_cs      = 0   # questions with contradiction_settings
        n_always  = 0   # questions always shown (show_condition = null)

        for qid_str, q_meta in questions.items():
            label = q_meta.get("label", qid_str)
            sc    = q_meta.get("show_condition")
            cs    = q_meta.get("contradiction_settings")

            # ── show_condition ─────────────────────────────────────────────
            if sc:
                n_sc += 1
                try:
                    should_see = eval_condition_vec(sc, df, label_to_q)
                    has_ans    = has_answer_vec(q_meta, df)
                    cond_str   = describe_condition(sc)

                    # Extra: answered but show_condition NOT met
                    extra_mask = (~should_see) & has_ans
                    for idx in df.index[extra_mask]:
                        pid = df.at[idx, pid_col]
                        violations.append({
                            "profile_id":        str(pid),
                            "type":              "show_condition_extra",
                            "question":          label,
                            "detail":            "answered but show_condition not met",
                            "condition_eval":    cond_str,
                            "condition_trigger": "",   # sc=False for this row → no triggered branch
                        })

                    # Missing: show_condition met but no answer
                    missing_mask = should_see & (~has_ans)
                    for idx in df.index[missing_mask]:
                        pid     = df.at[idx, pid_col]
                        row_df  = df.loc[[idx]]
                        try:
                            trigger = describe_trigger(sc, row_df, label_to_q)
                        except Exception:
                            trigger = ""
                        violations.append({
                            "profile_id":        str(pid),
                            "type":              "show_condition_missing",
                            "question":          label,
                            "detail":            "show_condition met but no answer recorded",
                            "condition_eval":    cond_str,
                            "condition_trigger": trigger,
                        })

                except Exception as exc:
                    logger.warning("show_condition eval error — %s: %s", label, exc)

            # ── contradiction_settings ─────────────────────────────────────
            if cs and cs.get("rules"):
                n_cs += 1
                try:
                    triggered  = eval_condition_vec(cs["rules"], df, label_to_q)
                    cond_str   = describe_condition(cs["rules"])

                    for idx in df.index[triggered]:
                        pid    = df.at[idx, pid_col]
                        row_df = df.loc[[idx]]
                        try:
                            trigger = describe_trigger(cs["rules"], row_df, label_to_q)
                        except Exception:
                            trigger = ""
                        violations.append({
                            "profile_id":        str(pid),
                            "type":              "contradiction",
                            "question":          label,
                            "detail":            "contradiction rule triggered",
                            "condition_eval":    cond_str,
                            "condition_trigger": trigger,
                        })

                except Exception as exc:
                    logger.warning("contradiction eval error — %s: %s", label, exc)

            # ── always shown (show_condition = null) — check missing data ──
            # sc is None means the question is always displayed to respondents.
            # Flag any respondent who has no answer for a trackable question.
            if sc is None and q_meta.get("answer_type") in _TRACKABLE_TYPES:
                raw_cols = q_meta.get("rawdata_columns") or []
                if not raw_cols:
                    continue  # no rawdata columns → not trackable, skip
                n_always += 1
                try:
                    has_ans      = has_answer_vec(q_meta, df)
                    missing_mask = ~has_ans
                    mandatory    = q_meta.get("mandatory", False)
                    for idx in df.index[missing_mask]:
                        pid = df.at[idx, pid_col]
                        violations.append({
                            "profile_id":        str(pid),
                            "type":              "always_shown_missing",
                            "question":          label,
                            "detail":            (
                                "mandatory question has no answer"
                                if mandatory
                                else "optional question has no answer"
                            ),
                            "condition_eval":    "always shown (no condition)",
                            "condition_trigger": "",
                        })
                except Exception as exc:
                    logger.warning("always_shown check error — %s: %s", label, exc)

        # ── Assemble report ────────────────────────────────────────────────
        flagged_ids  = {v["profile_id"] for v in violations}
        type_counts  = Counter(v["type"] for v in violations)

        report: dict = {
            "survey_id":         metadata.get("survey_id"),
            "total_respondents": len(df),
            "questions_checked": {
                "show_condition":         n_sc,
                "contradiction_settings": n_cs,
                "always_shown":           n_always,
            },
            "summary": {
                "flagged_count":          len(flagged_ids),
                "show_condition_extra":   type_counts.get("show_condition_extra", 0),
                "show_condition_missing": type_counts.get("show_condition_missing", 0),
                "always_shown_missing":   type_counts.get("always_shown_missing", 0),
                "contradiction":          type_counts.get("contradiction", 0),
            },
            "violations": violations,
        }

        report_path = quality_dir / "quality_report.json"
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )

        # ── CSV flat file ──────────────────────────────────────────────────
        csv_path = quality_dir / "flagged_profiles.csv"
        cols = ["profile_id", "type", "question", "detail", "condition_eval", "condition_trigger"]
        vdf  = pd.DataFrame(violations, columns=cols) if violations else pd.DataFrame(columns=cols)
        vdf.to_csv(csv_path, index=False, encoding="utf-8-sig")

        logger.info(
            "Quality check done — %d violation(s) across %d respondent(s)  →  %s",
            len(violations), len(flagged_ids), quality_dir,
        )

        context["quality_report"]       = report
        context["quality_dir"]          = str(quality_dir)
        context["quality_report_path"]  = str(report_path)
        context["quality_csv_path"]     = str(csv_path)
        return context
