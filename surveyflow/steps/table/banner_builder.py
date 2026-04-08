"""Build banner column definitions from datatable config."""
from __future__ import annotations

from dataclasses import dataclass
import pandas as pd


def _letter(i: int) -> str:
    """0→A, 1→B, …, 25→Z, 26→AA, …"""
    letters = []
    n = i + 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        letters.append(chr(65 + r))
    return "".join(reversed(letters))


@dataclass
class BannerColumn:
    group_label: str      # e.g. "Gender"  — used to group sig test pairs
    subgroup_label: str   # e.g. "Male"
    letter: str           # A, B, C … resets to A at each new group
    mask: pd.Series       # boolean mask into the rawdata DataFrame
    is_total: bool = False  # True → excluded from sig test


def build_banner(config: dict, df: pd.DataFrame) -> list[BannerColumn]:
    """Return one BannerColumn per banner subgroup defined in config."""
    columns: list[BannerColumn] = []

    for entry in config.get("banner", []):
        group_label = entry["label"]

        # ── Total ──────────────────────────────────────────────────────
        # Detect Total: no "groups" key AND no "question" key.
        # (Multi-condition cross-banners have "groups" but no top-level "question"
        # — they must NOT be treated as Total.)
        if "groups" not in entry and "question" not in entry:
            columns.append(BannerColumn(
                group_label=group_label,
                subgroup_label="Total",
                letter="",                              # no letter for Total
                mask=pd.Series(True, index=df.index),
                is_total=True,
            ))
            continue

        # ── Question-based breakdown ───────────────────────────────────
        for idx, grp in enumerate(entry.get("groups", [])):

            # ── Multi-condition group (Gender × Age, etc.) ─────────
            if "conditions" in grp:
                mask = pd.Series(True, index=df.index)
                for cond in grp["conditions"]:
                    cq = cond["question"]
                    if "value" in cond:
                        mask = mask & (df[cq] == cond["value"])
                    elif "values" in cond:
                        mask = mask & df[cq].isin(cond["values"])

            # ── Single-question group ──────────────────────────────
            else:
                q = entry["question"]
                if "value" in grp:
                    mask = df[q] == grp["value"]
                elif "values" in grp:
                    mask = df[q].isin(grp["values"])
                else:
                    mask = pd.Series(False, index=df.index)

            columns.append(BannerColumn(
                group_label=group_label,
                subgroup_label=grp["label"],
                letter=_letter(idx),                   # resets to A each group
                mask=mask,
                is_total=False,
            ))

    return columns
