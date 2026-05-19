"""
main.py — SurveyFlow runner (no Claude required)

Chỉnh các biến trong phần RUN CONFIG bên dưới rồi chạy:
  python main.py
"""

import json
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

SURVEY_ID   = 723380             # int  — survey ID từ QMe
SURVEY_NAME = "VN8976_Online_survey_UA_ ED_RTD_TEA" # str  — tên thư mục output

# Bật/tắt từng bước
FETCH_MCP   = False  # True  = fetch lại data từ QMe (ghi đè mcp/)
RUN_INGEST  = True  # True  = chạy lại ingestion (ghi đè data/)
RUN_TABLE   = True   # True  = chạy table → datatable.xlsx

# Version cho table (None = tự tăng: v1 → v2 → v3 …)
TABLE_VERSION = None   # hoặc đặt cứng, ví dụ: "v2"

# ── Table export mode ───────────────────────────────────────
# "xlsx"   = compute + render → datatable.xlsx (chậm hơn, có file output)
# "json"   = compute only    → table_results dict (nhanh hơn, không có file)
TABLE_EXPORT = "xlsx"

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
    from surveyflow import FetchStep, Pipeline, PipelineConfig

    output_dir    = OUTPUT_ROOT / SURVEY_NAME
    data_dir      = output_dir / "data"
    mcp_dir       = output_dir / "mcp"
    datatable_cfg = output_dir / "datatable" / "datatable.json"

    if TABLE_EXPORT not in ("xlsx", "json"):
        sys.exit(f"Error: TABLE_EXPORT phải là 'xlsx' hoặc 'json', hiện tại: {TABLE_EXPORT!r}")

    print("=" * 60)
    print(f"  SurveyFlow Runner")
    print(f"  Survey  : {SURVEY_NAME}  (id={SURVEY_ID})")
    print(f"  Mode    : fetch={FETCH_MODE}  export={TABLE_EXPORT}")
    print(f"  Steps   : fetch={FETCH_MCP}  ingest={RUN_INGEST}  table={RUN_TABLE}")
    print("=" * 60)

    # ── Step 1: Fetch ──────────────────────────────────────────────────────────
    if FETCH_MODE == "export":
        mcp_ready = (mcp_dir / "data_export.csv").exists()
        sentinel  = "data_export.csv"
    else:
        mcp_ready = mcp_dir.exists() and any(mcp_dir.glob("rows_page_*.json"))
        sentinel  = "rows_page_*.json"

    fetch_ctx = {}
    if FETCH_MCP or not mcp_ready:
        if not FETCH_MCP:
            print(f"\n[fetch] {sentinel} not found → fetching automatically …")
        fetch_ctx = FetchStep().run({
            "client":    _make_client(MCP_URL),
            "survey_id": SURVEY_ID,
            "mcp_dir":   str(mcp_dir),
            "mode":      FETCH_MODE,
            "date_from": DATE_FROM,
            "date_to":   DATE_TO,
        })
    else:
        print(f"\n[fetch] Skipped ({sentinel} exists)")

    # ── Step 2 + 3: Ingestion + Table via Pipeline ─────────────────────────────
    skip_ingestion = (data_dir / "rawdata.csv").exists() and not RUN_INGEST
    if skip_ingestion:
        print(f"\n[ingest] Skipped (rawdata.csv exists)")

    # Resolve ingestion inputs (from FetchStep context or disk)
    definition, rows_pages, export_df = None, None, None
    if not skip_ingestion:
        definition = fetch_ctx.get("definition") or json.loads(
            (mcp_dir / "definition.json").read_text(encoding="utf-8")
        )
        if FETCH_MODE == "export":
            from surveyflow.steps.ingestion.export_parser import parse_export_csv
            export_csv = pathlib.Path(
                fetch_ctx.get("export_csv") or str(mcp_dir / "data_export.csv")
            )
            export_df = parse_export_csv(export_csv)
        else:
            rows_pages = fetch_ctx.get("rows_pages") or [
                json.loads(p.read_text(encoding="utf-8"))
                for p in sorted(mcp_dir.glob("rows_page_*.json"))
            ]

    # Resolve table config
    has_table    = RUN_TABLE and datatable_cfg.exists()
    version      = TABLE_VERSION or _next_version(SURVEY_NAME)
    if RUN_TABLE and not datatable_cfg.exists():
        print(f"\n[table] datatable.json not found at {datatable_cfg} — skipping table")

    skip_render = (TABLE_EXPORT == "json")

    result = Pipeline(PipelineConfig(
        definition       = definition,
        rows_pages       = rows_pages,
        export_df        = export_df,
        output_dir       = str(output_dir),
        data_dir         = str(data_dir),
        skip_ingestion   = skip_ingestion,
        skip_render      = skip_render,
        version          = version,
        datatable_config = str(datatable_cfg) if has_table else None,
    )).run()

    print(f"\n--- Output ---")
    print(f"  rawdata  → {result['rawdata_path']}")
    print(f"  metadata → {result['metadata_path']}")
    if "datatable_path" in result:
        print(f"  datatable → {result['datatable_path']}")
    if "table_results" in result:
        import time as _time
        n_configs = len(result["table_results"])
        n_blocks  = sum(len(t["blocks"]) for t in result["table_results"])
        json_dir  = output_dir / version
        json_dir.mkdir(parents=True, exist_ok=True)
        json_path = json_dir / "table_results.json"
        _t0 = _time.perf_counter()
        try:
            import orjson
            json_path.write_bytes(
                orjson.dumps(result["table_results"], option=orjson.OPT_INDENT_2)
            )
            _engine = "orjson"
        except ImportError:
            json_path.write_text(
                json.dumps(result["table_results"], ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            _engine = "json"
        _elapsed = _time.perf_counter() - _t0
        print(f"  table_results → {n_configs} config(s), {n_blocks} block(s)  →  {json_path}  ({_engine}, {_elapsed:.2f}s)")
    print(f"  rows     : {result['rawdata'].shape[0]}")
    print("\nDone.")


if __name__ == "__main__":
    main()
