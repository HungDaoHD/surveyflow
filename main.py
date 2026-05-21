"""
main.py — SurveyFlow runner (no Claude required)

Chỉnh các biến trong phần RUN CONFIG bên dưới rồi chạy:
  python main.py
"""

import logging
import os
import pathlib
import sys

from dotenv import load_dotenv

logging.basicConfig(
    level  = logging.INFO,
    format = "%(asctime)s  %(message)s",
    datefmt= "%H:%M:%S",
)

load_dotenv()

from qme_auth import get_access_token


# ══════════════════════════════════════════════════════════
#  RUN CONFIG  ← chỉnh tại đây
# ══════════════════════════════════════════════════════════

# SURVEY_ID   = 723334             # int  — survey ID từ QMe
# SURVEY_NAME = "VN8966_DUC_Random" # str  — tên thư mục output

SURVEY_ID   = 723306             # int  — survey ID từ QMe
SURVEY_NAME = "VN8954-BHT Job Site 2026" # str  — tên thư mục output

# Bật/tắt từng bước
FETCH_MCP   = True  # True  = fetch lại data từ QMe (ghi đè mcp/)
RUN_INGEST  = True  # True  = chạy lại ingestion (ghi đè data/)
RUN_TABLE   = True   # True  = chạy table → datatable.xlsx

# Version cho table (None = tự tăng: v1 → v2 → v3 …)
TABLE_VERSION = None   # hoặc đặt cứng, ví dụ: "v2"

# ── Fetch mode ──────────────────────────────────────────────
# "rows"   = get_survey_definition + get_survey_rows
# "export" = get_survey_definition + prepare_survey_data_file
FETCH_MODE = "export"

# Date range — rows mode only
DATE_FROM = "2020-01-01"
DATE_TO   = "2030-12-31"

# ══════════════════════════════════════════════════════════

MCP_URL     = os.getenv("QME_MCP_BASE_URL", "https://retail.qand.me/api/mcp")
CLIENT_ID   = os.getenv("QME_CLIENT_ID", "")
OUTPUT_ROOT = pathlib.Path("output")

if not CLIENT_ID:
    sys.exit("Error: QME_CLIENT_ID not set. Add it to .env file.")


# ── MCP Client ─────────────────────────────────────────────────────────────────

from surveyflow import QMeClient


def _extract_http_error(exc: BaseException, cls: type):
    """Return the first instance of *cls* found in *exc* or its ExceptionGroup chain."""
    if isinstance(exc, cls):
        return exc
    if isinstance(exc, BaseExceptionGroup):
        for child in exc.exceptions:
            found = _extract_http_error(child, cls)
            if found is not None:
                return found
    return None


def _make_client(url: str) -> QMeClient:
    """Return a QMeClient with automatic 401-retry."""
    import httpx

    token = get_access_token()

    class _RetryClient(QMeClient):
        async def _call_async(self, tool_name, arguments):
            try:
                return await super()._call_async(tool_name, arguments)
            except BaseException as exc:
                http_err = _extract_http_error(exc, httpx.HTTPStatusError)
                if http_err is not None and http_err.response.status_code in (401, 403):
                    print("[auth] Token rejected (401) — acquiring new token …")
                    self.token = get_access_token()
                    return await super()._call_async(tool_name, arguments)
                raise

    return _RetryClient(url, token)


# ── Helpers ────────────────────────────────────────────────────────────────────

def _next_version(survey_name: str) -> str:
    existing = sorted((OUTPUT_ROOT / survey_name).glob("v*/datatable.xlsx"))
    return f"v{len(existing) + 1}"


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    from surveyflow import FetchStep
    from surveyflow.cli import main as cli_main

    output_dir    = OUTPUT_ROOT / SURVEY_NAME
    data_dir      = output_dir / "data"
    mcp_dir       = output_dir / "mcp"
    datatable_cfg = output_dir / "datatable" / "datatable.json"

    print("=" * 60)
    print(f"  SurveyFlow Runner")
    print(f"  Survey  : {SURVEY_NAME}  (id={SURVEY_ID})")
    print(f"  Mode    : fetch={FETCH_MODE}")
    print(f"  Steps   : fetch={FETCH_MCP}  ingest={RUN_INGEST}  table={RUN_TABLE}")
    print("=" * 60)

    # ── Step 1: Fetch ──────────────────────────────────────────────────────────
    if FETCH_MODE == "export":
        mcp_ready = (mcp_dir / "data_export.csv").exists()
        sentinel  = "data_export.csv"
    else:
        mcp_ready = mcp_dir.exists() and any(mcp_dir.glob("rows_page_*.json"))
        sentinel  = "rows_page_*.json"

    if FETCH_MCP or not mcp_ready:
        if not FETCH_MCP:
            print(f"\n[fetch] {sentinel} not found → fetching automatically …")
        FetchStep().run({
            "client":    _make_client(MCP_URL),
            "survey_id": SURVEY_ID,
            "mcp_dir":   str(mcp_dir),
            "mode":      FETCH_MODE,
            "date_from": DATE_FROM,
            "date_to":   DATE_TO,
        })
    else:
        print(f"\n[fetch] Skipped ({sentinel} exists)")

    # ── Step 2 + 3: Ingestion + Table via CLI ─────────────────────────────────
    skip_ingestion = (data_dir / "rawdata.csv").exists() and not RUN_INGEST
    if skip_ingestion:
        print(f"\n[ingest] Skipped (rawdata.csv exists)")

    version   = TABLE_VERSION or _next_version(SURVEY_NAME)
    has_table = RUN_TABLE and datatable_cfg.exists()
    if RUN_TABLE and not datatable_cfg.exists():
        print(f"\n[table] datatable.json not found at {datatable_cfg} — skipping table")

    cli_argv = ["--output-dir", str(output_dir), "--version", version,
                "--profile-status", "approved,pending"]
    if not skip_ingestion:
        cli_argv += ["--mcp-dir", str(mcp_dir), "--force-ingestion"]
        if FETCH_MODE == "export":
            cli_argv += ["--export-csv", str(mcp_dir / "data_export.csv")]
    if has_table:
        cli_argv += ["--datatable-config", str(datatable_cfg)]

    cli_main(cli_argv)
    print("\nDone.")


if __name__ == "__main__":
    main()
