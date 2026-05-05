"""
main.py — SurveyFlow runner (no Claude required)

Chỉnh các biến trong phần RUN CONFIG bên dưới rồi chạy:
  python main.py
"""

import json
import os
import pathlib
import sys

import requests
from dotenv import load_dotenv

load_dotenv()

# ══════════════════════════════════════════════════════════
#  RUN CONFIG  ← chỉnh tại đây
# ══════════════════════════════════════════════════════════

SURVEY_ID   = 723061                    # int  — survey ID từ QMe
SURVEY_NAME = "VN8931_Baby_Food_Package"  # str  — tên thư mục output

# Bật/tắt từng bước
FETCH_MCP   = False   # True  = fetch lại data từ QMe (ghi đè mcp/)
RUN_INGEST  = False   # True  = chạy lại ingestion (ghi đè data/)
RUN_TABLE   = True    # True  = chạy table → datatable.xlsx

# Version cho table (None = tự tăng: v1 → v2 → v3 …)
TABLE_VERSION = None   # hoặc đặt cứng, ví dụ: "v2"

# ══════════════════════════════════════════════════════════

MCP_URL     = os.getenv("QME_MCP_URL", "https://retail.qand.me/api/mcp")
API_KEY     = os.getenv("QME_API_KEY", "")
OUTPUT_ROOT = pathlib.Path("output")

if not API_KEY:
    sys.exit("Error: QME_API_KEY not set. Add it to .env file.")


# ── MCP Client ─────────────────────────────────────────────────────────────────

class MCPClient:
    """Thin HTTP wrapper for the QMe MCP API."""

    def __init__(self, url: str, key: str):
        self.url     = url
        self.params  = {"key": key}
        self._req_id = 0

    def _call(self, tool_name: str, arguments: dict) -> dict:
        self._req_id += 1
        payload = {
            "jsonrpc": "2.0",
            "method":  "tools/call",
            "params":  {"name": tool_name, "arguments": arguments},
            "id":      self._req_id,
        }
        resp = requests.post(self.url, params=self.params, json=payload, timeout=60)
        resp.raise_for_status()
        data = resp.json()
        if "error" in data:
            raise RuntimeError(f"MCP error: {data['error']}")
        result = data.get("result", {})
        if isinstance(result, dict) and "content" in result:
            for item in result["content"]:
                if item.get("type") == "text":
                    return json.loads(item["text"])
        return result

    def get_survey_definition(self, survey_id: int) -> dict:
        return self._call("get_survey_definition", {"survey_id": survey_id})

    def get_survey_rows(self, survey_id: int, offset: int = 0, limit: int = 200) -> dict:
        return self._call("get_survey_rows", {
            "survey_id": survey_id,
            "format":    "code",
            "offset":    offset,
            "limit":     limit,
        })


# ── Step helpers ───────────────────────────────────────────────────────────────

def fetch_mcp(client: MCPClient, survey_id: int, survey_name: str) -> pathlib.Path:
    """Fetch definition + all row pages → output/<name>/mcp/"""
    mcp_dir = OUTPUT_ROOT / survey_name / "mcp"
    mcp_dir.mkdir(parents=True, exist_ok=True)

    print("\n[fetch] Fetching definition …")
    definition = client.get_survey_definition(survey_id)
    (mcp_dir / "definition.json").write_text(
        json.dumps(definition, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print("  → definition.json saved")

    print("[fetch] Fetching rows …")
    page, offset = 1, 0
    while True:
        rows_data = client.get_survey_rows(survey_id, offset=offset)
        (mcp_dir / f"rows_page_{page}.json").write_text(
            json.dumps(rows_data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        count = len(rows_data.get("rows", []))
        print(f"  → page {page}: {count} rows")
        if not rows_data.get("pagination", {}).get("has_more"):
            break
        offset += count
        page   += 1

    print(f"  Done → {mcp_dir}")
    return mcp_dir


def run_ingestion(survey_name: str) -> pathlib.Path:
    """Run surveyflow ingestion → rawdata.csv + metadata.json"""
    from surveyflow import Pipeline, PipelineConfig

    output_dir = OUTPUT_ROOT / survey_name
    data_dir   = output_dir / "data"
    mcp_dir    = output_dir / "mcp"

    definition = json.loads((mcp_dir / "definition.json").read_text(encoding="utf-8"))
    rows_pages = [
        json.loads(p.read_text(encoding="utf-8"))
        for p in sorted(mcp_dir.glob("rows_page_*.json"))
    ]

    print("\n[ingest] Running ingestion …")
    result = Pipeline(PipelineConfig(
        definition = definition,
        rows_pages = rows_pages,
        output_dir = str(output_dir),
        data_dir   = str(data_dir),
    )).run()

    print(f"  rawdata  → {result['rawdata_path']}")
    print(f"  metadata → {result['metadata_path']}")
    print(f"  rows     : {result['rawdata'].shape[0]}")
    return data_dir


def run_table(survey_name: str, version: str) -> pathlib.Path | None:
    """Run surveyflow table step → datatable.xlsx"""
    from surveyflow import Pipeline, PipelineConfig

    output_dir    = OUTPUT_ROOT / survey_name
    data_dir      = output_dir / "data"
    datatable_cfg = output_dir / "datatable" / "datatable.json"

    if not datatable_cfg.exists():
        print(f"\n[table] datatable.json not found at {datatable_cfg}")
        print("  Create it first using datatable-editor.html, then set RUN_TABLE = True.")
        return None

    print(f"\n[table] Running table ({version}) …")
    result = Pipeline(PipelineConfig(
        output_dir       = str(output_dir),
        data_dir         = str(data_dir),
        skip_ingestion   = True,
        version          = version,
        datatable_config = str(datatable_cfg),
    )).run()

    path = pathlib.Path(result["datatable_path"])
    print(f"  datatable → {path}")
    return path


def _next_version(survey_name: str) -> str:
    existing = sorted((OUTPUT_ROOT / survey_name).glob("v*/datatable.xlsx"))
    return f"v{len(existing) + 1}"


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print("=" * 55)
    print(f"  SurveyFlow Runner")
    print(f"  Survey  : {SURVEY_NAME}  (id={SURVEY_ID})")
    print(f"  Steps   : fetch={FETCH_MCP}  ingest={RUN_INGEST}  table={RUN_TABLE}")
    print("=" * 55)

    client = MCPClient(MCP_URL, API_KEY)

    if FETCH_MCP:
        fetch_mcp(client, SURVEY_ID, SURVEY_NAME)
    else:
        mcp_dir = OUTPUT_ROOT / SURVEY_NAME / "mcp"
        if mcp_dir.exists():
            print(f"\n[fetch] Skipped (data exists at {mcp_dir})")
        else:
            print(f"\n[fetch] mcp/ not found → fetching automatically …")
            fetch_mcp(client, SURVEY_ID, SURVEY_NAME)

    if RUN_INGEST:
        run_ingestion(SURVEY_NAME)
    else:
        rawdata = OUTPUT_ROOT / SURVEY_NAME / "data" / "rawdata.csv"
        if rawdata.exists():
            print(f"\n[ingest] Skipped (rawdata exists at {rawdata})")
        else:
            print(f"\n[ingest] rawdata.csv not found → running ingestion automatically …")
            run_ingestion(SURVEY_NAME)

    if RUN_TABLE:
        version = TABLE_VERSION or _next_version(SURVEY_NAME)
        run_table(SURVEY_NAME, version)
    else:
        print(f"\n[table] Skipped (RUN_TABLE = False)")

    print("\nDone.")


if __name__ == "__main__":
    main()
