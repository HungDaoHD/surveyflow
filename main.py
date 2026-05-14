"""
main.py — SurveyFlow runner (no Claude required)

Chỉnh các biến trong phần RUN CONFIG bên dưới rồi chạy:
  python main.py
"""

import base64
import hashlib
import json
import os
import pathlib
import re
import secrets
import sys
from urllib.parse import parse_qs, urlencode, urlparse

import requests
from dotenv import load_dotenv

load_dotenv()


# ══════════════════════════════════════════════════════════
#  RUN CONFIG  ← chỉnh tại đây
# ══════════════════════════════════════════════════════════

SURVEY_ID   = 723334                    # int  — survey ID từ QMe
SURVEY_NAME = "VN8966_DUC"  # str  — tên thư mục output

# Bật/tắt từng bước
FETCH_MCP   = True   # True  = fetch lại data từ QMe (ghi đè mcp/)
RUN_INGEST  = True   # True  = chạy lại ingestion (ghi đè data/)
RUN_TABLE   = True    # True  = chạy table → datatable.xlsx

# Version cho table (None = tự tăng: v1 → v2 → v3 …)
TABLE_VERSION = None   # hoặc đặt cứng, ví dụ: "v2"

# Date range for fetching rows (get_survey_rows requires both)
DATE_FROM = "2020-01-01"   # wide default — covers all survey history
DATE_TO   = "2030-12-31"   # far future — fetches everything up to now

# ══════════════════════════════════════════════════════════

MCP_URL       = os.getenv("QME_MCP_BASE_URL", "https://retail.qand.me/api/mcp")
CLIENT_ID     = os.getenv("QME_CLIENT_ID", "")
CLIENT_SECRET = os.getenv("QME_CLIENT_SECRET", "")
QME_USERNAME  = os.getenv("QME_USER_NAME", "")
QME_PASSWORD  = os.getenv("QME_USER_PASS", "")
OUTPUT_ROOT   = pathlib.Path("output")

if not CLIENT_ID or not CLIENT_SECRET:
    sys.exit("Error: QME_CLIENT_ID / QME_CLIENT_SECRET not set. Add them to .env file.")

# ── OAuth2 Authorization Code + PKCE ──────────────────────────────────────────

_QME_BASE       = "https://retail.qand.me"
_AUTH_ENDPOINT  = f"{_QME_BASE}/api/mcp/oauth/authorize"
_TOKEN_ENDPOINT = f"{_QME_BASE}/api/mcp/oauth/token"
_SCOPE          = "mcp mcp:read"
_REDIRECT_URI   = "https://claude.ai/api/mcp/auth_callback"
_RESOURCE       = f"{_QME_BASE}/api/mcp"   # RFC 8707 resource indicator
_TOKEN_FILE     = pathlib.Path(__file__).parent / ".mcp_token.json"


def _pkce_pair():
    verifier  = base64.urlsafe_b64encode(os.urandom(32)).rstrip(b"=").decode()
    challenge = base64.urlsafe_b64encode(
        hashlib.sha256(verifier.encode()).digest()
    ).rstrip(b"=").decode()
    return verifier, challenge


def _extract_csrf_meta(html: str) -> str:
    """Extract CSRF token from <meta name="csrf-token" content="...">."""
    for line in html.splitlines():
        if 'name="csrf-token"' in line and 'content="' in line:
            idx = line.index('content="') + 9
            return line[idx : line.index('"', idx)]
    return ""


def _extract_form_token(html: str) -> str:
    """Extract hidden _token field from a Laravel form."""
    m = re.search(r'<input[^>]+name=["\']_token["\'][^>]+value=["\']([^"\']+)["\']', html)
    if not m:
        m = re.search(r'<input[^>]+value=["\']([^"\']+)["\'][^>]+name=["\']_token["\']', html)
    return m.group(1) if m else ""


def _extract_code(urls: list) -> tuple:
    """Return (code, redirect_uri_base) from first URL that has a code param."""
    for url in urls:
        if not url:
            continue
        parsed = urlparse(url)
        params = parse_qs(parsed.query)
        code = (params.get("code") or [None])[0]
        if code:
            redirect_uri = f"{parsed.scheme}://{parsed.netloc}{parsed.path}"
            return code, redirect_uri
    return "", ""


def _save_tokens(tokens: dict):
    _TOKEN_FILE.write_text(json.dumps(tokens, indent=2), encoding="utf-8")


def _load_tokens():
    if _TOKEN_FILE.exists():
        try:
            return json.loads(_TOKEN_FILE.read_text(encoding="utf-8"))
        except Exception:
            return None
    return None


def _refresh_token(refresh_tok: str):
    r = requests.post(_TOKEN_ENDPOINT, data={
        "grant_type":    "refresh_token",
        "refresh_token": refresh_tok,
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "resource":      _RESOURCE,
    }, timeout=30)
    if r.status_code == 200:
        return r.json()
    return None



def _authorize_headless() -> dict:
    """Headless OAuth2 + PKCE flow — mirrors apdp qme_auth.py exactly."""
    verifier, challenge = _pkce_pair()
    state = secrets.token_urlsafe(32)

    auth_url = _AUTH_ENDPOINT + "?" + urlencode({
        "response_type":         "code",
        "client_id":             CLIENT_ID,
        "redirect_uri":          _REDIRECT_URI,
        "state":                 state,
        "code_challenge":        challenge,
        "code_challenge_method": "S256",
        "scope":                 _SCOPE,
        "resource":              _RESOURCE,
    })

    pw_hash = hashlib.sha256(QME_PASSWORD.encode()).hexdigest()
    session = requests.Session()

    # ── Step 1: GET /authorize → QMe sets OAuth context, redirects to login2 ──
    r0 = session.get(auth_url, allow_redirects=True, timeout=30)
    csrf = _extract_csrf_meta(r0.text) or _extract_form_token(r0.text)
    if not csrf:
        sys.exit(f"[auth] No CSRF token on login page (url={r0.url})")

    # ── Step 2: POST DoLogin2 (form data + isShareClient) ──────────────────────
    r1 = session.post(
        f"{_QME_BASE}/Admin/DoLogin2",
        data={
            "email":         QME_USERNAME,
            "password":      pw_hash,
            "_token":        csrf,
            "login2":        "1",
            "countrycode":   "",
            "isShareClient": "1",
        },
        headers={
            "X-CSRF-TOKEN":     csrf,
            "X-Requested-With": "XMLHttpRequest",
            "Referer":          f"{_QME_BASE}/login2",
        },
        timeout=15,
    )
    d1 = r1.json()
    if d1.get("result") != 1:
        sys.exit(f"[auth] Login failed: {d1.get('msg') or d1.get('message')}")

    # Extract ref from redirect URL
    login_redirect = d1.get("redirect", "")
    ref = d1.get("ref", "")
    if not ref and "ref=" in login_redirect:
        ref = parse_qs(urlparse(login_redirect).query).get("ref", [""])[0]

    # ── Step 3: GET loginverify to arm OTP session + get verify CSRF ───────────
    csrf_verify = ""
    if ref:
        r_lv = session.get(f"{_QME_BASE}/loginverify", params={"ref": ref}, timeout=15)
        csrf_verify = _extract_csrf_meta(r_lv.text) or _extract_form_token(r_lv.text)

    # ── Step 4: Ask for OTP ────────────────────────────────────────────────────
    print("\n[auth] A verification code has been sent to your email.")
    otp = input("[auth] Enter the OTP code: ").strip()

    # ── Step 5: POST DoVerify (form-encoded, XSRF from cookie) ─────────────────
    xsrf   = session.cookies.get("XSRF-TOKEN", "")
    post_data: dict = {"verifyCode": otp, "ref": ref}
    if csrf_verify:
        post_data["_token"] = csrf_verify

    r2 = session.post(
        f"{_QME_BASE}/Admin/DoVerify",
        data=post_data,
        headers={
            "X-Requested-With": "XMLHttpRequest",
            "X-XSRF-TOKEN":     xsrf,
            "Referer":          f"{_QME_BASE}/loginverify?ref={ref}",
        },
        allow_redirects=False,
        timeout=15,
    )
    try:
        d2 = r2.json()
    except Exception:
        d2 = {}
    if not d2 or d2.get("result") != 1:
        sys.exit(f"[auth] OTP failed: {d2.get('msg') or d2.get('message') or r2.text[:100]}")

    # ── Step 6: Follow DoVerify redirect chain → find OAuth code ──────────────
    # DoVerify returns either Location header or data.redirect pointing back to /authorize.
    # The authorize endpoint may itself redirect several times before reaching the callback.
    next_url = r2.headers.get("location") or d2.get("redirect", "")
    code, redirect_uri = _extract_code([next_url])

    cur_url = next_url
    for hop in range(6):
        if code or not cur_url:
            break
        if not cur_url.startswith("http"):
            cur_url = _QME_BASE + cur_url
        print(f"[auth] redirect hop {hop}: {cur_url[:100]}")
        r_hop = session.get(cur_url, allow_redirects=False, timeout=15)
        print(f"[auth]   → {r_hop.status_code}  Location: {r_hop.headers.get('location', '(none)')[:120]}")
        cur_url = r_hop.headers.get("location", "")
        code, redirect_uri = _extract_code([cur_url])

    if not code:
        sys.exit(f"[auth] OTP verified but no OAuth code received. Last URL: {cur_url!r}")

    # Use the actual redirect_uri captured from the callback (may differ from configured)
    effective_redirect = redirect_uri or _REDIRECT_URI

    # ── Step 7: Exchange code for tokens ───────────────────────────────────────
    r_tok = requests.post(_TOKEN_ENDPOINT, data={
        "grant_type":    "authorization_code",
        "code":          code,
        "redirect_uri":  effective_redirect,
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "code_verifier": verifier,
    }, timeout=30)
    r_tok.raise_for_status()
    tok_data = r_tok.json()
    print(f"[auth] Token exchange OK — type={tok_data.get('token_type')} "
          f"expires_in={tok_data.get('expires_in')} "
          f"has_refresh={bool(tok_data.get('refresh_token'))}")
    return tok_data


def get_access_token() -> str:
    """Return a valid Bearer access token, refreshing or re-authorizing as needed."""

    # 1. Manual token override via env (for debugging)
    manual = os.getenv("QME_ACCESS_TOKEN", "").strip()
    if manual:
        print("[auth] Using QME_ACCESS_TOKEN from environment.")
        return manual

    # 2. Try saved refresh token
    tokens = _load_tokens()
    if tokens and tokens.get("refresh_token"):
        new = _refresh_token(tokens["refresh_token"])
        if new and new.get("access_token"):
            _save_tokens({**tokens, **new})
            return new["access_token"]

    # 3. Full headless login flow (requires QME_USER_NAME + QME_USER_PASS)
    if not QME_USERNAME or not QME_PASSWORD:
        sys.exit(
            "Error: QME_USER_NAME / QME_USER_PASS not set.\n"
            "Add them to .env file to enable headless login."
        )
    tokens = _authorize_headless()
    _save_tokens(tokens)
    print("[auth] Token saved to", _TOKEN_FILE)
    return tokens["access_token"]


# ── MCP Client ─────────────────────────────────────────────────────────────────

import asyncio
from mcp import ClientSession
from mcp.client.streamable_http import streamablehttp_client


class MCPClient:
    """QMe MCP client using the official MCP SDK (streamable HTTP transport)."""

    def __init__(self, url: str):
        self.url    = url
        self._token = None

    def _get_token(self) -> str:
        if not self._token:
            self._token = get_access_token()
        return self._token

    def _call(self, tool_name: str, arguments: dict) -> dict:
        return asyncio.run(self._call_async(tool_name, arguments))

    async def _call_async(self, tool_name: str, arguments: dict) -> dict:
        import httpx
        headers = {"Authorization": f"Bearer {self._get_token()}"}
        try:
            async with streamablehttp_client(self.url, headers=headers) as (r, w, _):
                async with ClientSession(r, w) as session:
                    await session.initialize()
                    result = await session.call_tool(tool_name, arguments)
        except* httpx.HTTPStatusError as eg:
            exc = eg.exceptions[0]
            if exc.response.status_code in (401, 403):
                # Token rejected — clear in-memory token and retry via refresh/re-login
                # Do NOT delete the token file here: it holds the refresh_token we need
                print("[auth] Token rejected (401) — acquiring new token …")
                self._token = None
                self._token = get_access_token()
                headers = {"Authorization": f"Bearer {self._token}"}
                async with streamablehttp_client(self.url, headers=headers) as (r, w, _):
                    async with ClientSession(r, w) as session:
                        await session.initialize()
                        result = await session.call_tool(tool_name, arguments)
            else:
                raise

        if result.content and hasattr(result.content[0], "text"):
            try:
                return json.loads(result.content[0].text)
            except Exception:
                return result.content[0].text
        return result

    def list_tools(self) -> list:
        return asyncio.run(self._list_tools_async())

    async def _list_tools_async(self) -> list:
        headers = {"Authorization": f"Bearer {self._get_token()}"}
        async with streamablehttp_client(self.url, headers=headers) as (r, w, _):
            async with ClientSession(r, w) as session:
                await session.initialize()
                result = await session.list_tools()
                return result.tools

    def get_survey_definition(self, survey_id: int) -> dict:
        return self._call("get_survey_definition", {"survey_id": survey_id})

    def get_survey_rows(self, survey_id: int, offset: int = 0, limit: int = 200,
                        date_from: str = DATE_FROM, date_to: str = DATE_TO,
                        profile_status: list | None = None) -> dict:
        args: dict = {
            "survey_id": survey_id,
            "date_from": date_from,
            "date_to":   date_to,
            "format":    "code",
            "offset":    offset,
            "limit":     limit,
        }
        if profile_status is not None:
            args["profile_status"] = profile_status
        return self._call("get_survey_rows", args)


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

    client = MCPClient(MCP_URL)

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
