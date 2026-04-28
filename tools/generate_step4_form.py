#!/usr/bin/env python3
"""
SurveyFlow — Step 4 Form Generator
====================================
Reads metadata.json → generates an interactive HTML form for Claude to show in the preview panel.
User fills the form (table type / banner / stub) and submits.
Claude reads window.__result via preview_eval.

Usage:
    python tools/generate_step4_form.py <metadata_path> [output_html_path] [survey_name]

Example:
    python tools/generate_step4_form.py "output/VN8932/data/metadata.json" "output/VN8932/step4_form.html" "VN8932"

window.__result structure after submit:
{
    "submitted": true,
    "table_type": 4,          // 1=Count, 2=Pct, 3=Pct+Sig, 4=All
    "banner": ["Q1", "Q3"],   // question labels selected for banner
    "stub_mode": "all",       // "all" | "select"
    "stub": ["Q1","Q5","Q8"]  // only when stub_mode == "select"
}
"""

import json
import sys
import pathlib
import html as _html


# ─── helpers ──────────────────────────────────────────────────────────────────

def get_label(obj, langs=("vi", "en")):
    if isinstance(obj, str):
        return obj
    if isinstance(obj, dict):
        for lang in langs:
            if lang in obj and obj[lang]:
                return obj[lang]
        return next(iter(obj.values()), "") if obj else ""
    return str(obj) if obj is not None else ""


def esc(s):
    return _html.escape(str(s))


# ─── main ─────────────────────────────────────────────────────────────────────

def generate_form(metadata_path: str, output_path: str, survey_name: str = ""):
    meta = json.load(open(metadata_path, encoding="utf-8"))

    # questions can be a list OR a dict keyed by question_id
    raw = meta.get("questions", {})
    if isinstance(raw, dict):
        questions = sorted(raw.values(), key=lambda q: q.get("position", 9999))
    else:
        questions = raw

    BANNER_TYPES = {"SA", "MA"}
    STUB_TYPES   = {"SA", "MA", "Matrix_SA", "Matrix_MA", "Matrix_NUM", "Matrix_NUM"}

    # answer_type field name may vary
    def qtype(q):
        return q.get("answer_type") or q.get("type") or ""

    # title field: question_i18n or title or label
    def qtitle(q):
        t = q.get("question_i18n") or q.get("title") or q.get("label", "")
        return get_label(t)

    banner_qs = [q for q in questions if qtype(q) in BANNER_TYPES]
    stub_qs   = [q for q in questions if qtype(q) in STUB_TYPES]

    def q_to_dict(q):
        return {"label": q["label"], "title": qtitle(q), "type": qtype(q)}

    banner_js = json.dumps([q_to_dict(q) for q in banner_qs], ensure_ascii=False)
    stub_js   = json.dumps([q_to_dict(q) for q in stub_qs],   ensure_ascii=False)
    name_js   = json.dumps(survey_name)

    html = f"""<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Step 4 — Cấu hình bảng</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;font-size:14px;background:#f0f2f5;color:#1a1a2e;min-height:100vh}}

.page{{max-width:780px;margin:0 auto;padding:24px 16px 60px}}

/* Header */
.page-header{{background:#16213e;color:white;border-radius:12px;padding:18px 24px;margin-bottom:20px;display:flex;align-items:center;gap:14px}}
.page-header .icon{{font-size:28px}}
.page-header h1{{font-size:17px;font-weight:700}}
.page-header p{{font-size:12px;color:#8ab4f8;margin-top:3px}}

/* Steps indicator */
.steps{{display:flex;gap:6px;margin-bottom:20px}}
.step-dot{{flex:1;height:4px;border-radius:2px;background:#e0e4ea}}
.step-dot.done{{background:#1a73e8}}
.step-dot.active{{background:#1a73e8;opacity:.4}}

/* Section card */
.card{{background:white;border:1px solid #e0e4ea;border-radius:10px;margin-bottom:16px;overflow:hidden}}
.card-header{{padding:14px 18px;border-bottom:1px solid #e0e4ea;display:flex;align-items:center;gap:10px}}
.card-header .step-num{{width:26px;height:26px;border-radius:50%;background:#1a73e8;color:white;font-size:13px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0}}
.card-header h2{{font-size:14px;font-weight:700;color:#222;flex:1}}
.card-header .sub{{font-size:12px;color:#999}}
.card-body{{padding:16px 18px}}

/* Table type options */
.type-grid{{display:grid;grid-template-columns:1fr 1fr;gap:10px}}
.type-option{{border:2px solid #e0e4ea;border-radius:8px;padding:14px;cursor:pointer;transition:all .15s;position:relative}}
.type-option:hover{{border-color:#1a73e8;background:#fafbff}}
.type-option.selected{{border-color:#1a73e8;background:#eef4ff}}
.type-option input{{position:absolute;opacity:0;width:0;height:0}}
.type-option .to-icon{{font-size:22px;margin-bottom:6px}}
.type-option .to-name{{font-weight:700;font-size:13px;color:#222}}
.type-option .to-desc{{font-size:11px;color:#888;margin-top:3px;line-height:1.4}}
.type-option.selected .to-name{{color:#1a73e8}}
.type-option .check{{position:absolute;top:10px;right:10px;width:18px;height:18px;border-radius:50%;border:2px solid #ddd;display:flex;align-items:center;justify-content:center;font-size:10px;color:white}}
.type-option.selected .check{{background:#1a73e8;border-color:#1a73e8}}
.type-option.selected .check::after{{content:'✓'}}

/* Question list */
.q-search{{width:100%;padding:8px 12px;border:1px solid #dde1e7;border-radius:6px;font-size:13px;margin-bottom:10px;outline:none}}
.q-search:focus{{border-color:#1a73e8}}
.q-list{{max-height:260px;overflow-y:auto;border:1px solid #e0e4ea;border-radius:8px}}
.q-item{{display:flex;align-items:center;gap:10px;padding:9px 12px;border-bottom:1px solid #f0f2f5;cursor:pointer;transition:background .1s}}
.q-item:last-child{{border-bottom:none}}
.q-item:hover{{background:#f8f9fa}}
.q-item.checked{{background:#eef4ff}}
.q-item input[type=checkbox]{{width:15px;height:15px;accent-color:#1a73e8;flex-shrink:0;cursor:pointer}}
.q-item .q-label{{font-size:12px;font-weight:700;color:#1a73e8;min-width:45px;flex-shrink:0}}
.q-item .q-title{{font-size:13px;color:#333;flex:1}}
.q-item .q-type{{font-size:10px;background:#f0f2f5;color:#888;padding:2px 6px;border-radius:4px;flex-shrink:0}}
.q-select-all{{padding:6px 12px;font-size:12px;color:#1a73e8;cursor:pointer;display:flex;align-items:center;gap:6px;border-bottom:1px solid #e0e4ea;background:#fafbff}}
.q-select-all:hover{{background:#eef4ff}}

/* Stub mode */
.stub-mode{{display:flex;gap:10px;margin-bottom:14px}}
.mode-btn{{flex:1;padding:10px;border:2px solid #e0e4ea;border-radius:8px;cursor:pointer;text-align:center;transition:all .15s;background:white}}
.mode-btn:hover{{border-color:#1a73e8;background:#fafbff}}
.mode-btn.active{{border-color:#1a73e8;background:#eef4ff}}
.mode-btn .mb-icon{{font-size:20px;margin-bottom:4px}}
.mode-btn .mb-name{{font-size:13px;font-weight:700;color:#333}}
.mode-btn.active .mb-name{{color:#1a73e8}}
.mode-btn .mb-desc{{font-size:11px;color:#888;margin-top:2px}}

/* Selected summary */
.summary-bar{{background:#f8f9fa;border:1px solid #e0e4ea;border-radius:8px;padding:10px 14px;margin-top:10px;font-size:12px;color:#555;display:none}}
.summary-bar.visible{{display:block}}
.summary-bar strong{{color:#1a73e8}}

/* Submit */
.submit-section{{margin-top:24px;text-align:center}}
.btn-submit{{background:#1a73e8;color:white;border:none;padding:13px 40px;border-radius:8px;font-size:15px;font-weight:700;cursor:pointer;transition:background .15s;box-shadow:0 2px 8px rgba(26,115,232,.3)}}
.btn-submit:hover{{background:#1557b0}}
.btn-submit:disabled{{background:#ccc;cursor:not-allowed;box-shadow:none}}
.required-hint{{font-size:12px;color:#e74c3c;margin-top:8px;display:none}}
.required-hint.show{{display:block}}

/* Success state */
.success{{display:none;text-align:center;padding:40px 20px}}
.success.show{{display:block}}
.success .s-icon{{font-size:64px;margin-bottom:16px}}
.success h2{{font-size:20px;font-weight:700;color:#1e8449;margin-bottom:8px}}
.success p{{color:#666;font-size:14px;line-height:1.6}}
.success .s-summary{{background:#f0faf4;border:1px solid #a9dfbf;border-radius:8px;padding:14px 18px;margin-top:18px;text-align:left;font-size:13px;line-height:1.8;color:#333}}

/* Scrollbar */
::-webkit-scrollbar{{width:5px}}::-webkit-scrollbar-track{{background:transparent}}::-webkit-scrollbar-thumb{{background:#ccc;border-radius:3px}}
</style>
</head>
<body>
<div class="page">

  <!-- Header -->
  <div class="page-header">
    <div class="icon">📊</div>
    <div>
      <h1>Cấu hình Data Table</h1>
      <p id="surveyNameLabel">SurveyFlow — Step 4</p>
    </div>
  </div>

  <!-- Progress -->
  <div class="steps">
    <div class="step-dot done"></div>
    <div class="step-dot done"></div>
    <div class="step-dot done"></div>
    <div class="step-dot active"></div>
  </div>

  <!-- Main form -->
  <div id="formArea">

    <!-- Section 1: Table type -->
    <div class="card">
      <div class="card-header">
        <div class="step-num">1</div>
        <h2>Loại bảng</h2>
        <span class="sub">Chọn sheets sẽ xuất ra</span>
      </div>
      <div class="card-body">
        <div class="type-grid">
          <label class="type-option" id="opt1">
            <input type="radio" name="table_type" value="1" onchange="onTypeChange()">
            <div class="check"></div>
            <div class="to-icon">🔢</div>
            <div class="to-name">Count only</div>
            <div class="to-desc">Chỉ sheet số lượng</div>
          </label>
          <label class="type-option" id="opt2">
            <input type="radio" name="table_type" value="2" onchange="onTypeChange()">
            <div class="check"></div>
            <div class="to-icon">📐</div>
            <div class="to-name">Percentage only</div>
            <div class="to-desc">Chỉ sheet phần trăm</div>
          </label>
          <label class="type-option" id="opt3">
            <input type="radio" name="table_type" value="3" onchange="onTypeChange()">
            <div class="check"></div>
            <div class="to-icon">🔬</div>
            <div class="to-name">Percentage + Sig test</div>
            <div class="to-desc">Pct và kiểm định sig (90% & 95%)</div>
          </label>
          <label class="type-option" id="opt4">
            <input type="radio" name="table_type" value="4" onchange="onTypeChange()">
            <div class="check"></div>
            <div class="to-icon">✨</div>
            <div class="to-name">Tất cả</div>
            <div class="to-desc">Count + Pct + Sig test</div>
          </label>
        </div>
      </div>
    </div>

    <!-- Section 2: Banner -->
    <div class="card">
      <div class="card-header">
        <div class="step-num">2</div>
        <h2>Banner (header cột)</h2>
        <span class="sub">Total luôn được thêm tự động</span>
      </div>
      <div class="card-body">
        <input type="text" class="q-search" placeholder="🔍 Tìm câu hỏi..." oninput="filterList('banner', this.value)">
        <div class="q-list" id="bannerList">
          <div class="q-select-all" onclick="toggleAll('banner')">
            <input type="checkbox" id="bannerAllChk" onclick="event.stopPropagation()"> Chọn tất cả
          </div>
          <!-- populated by JS -->
        </div>
        <div class="summary-bar" id="bannerSummary"></div>
      </div>
    </div>

    <!-- Section 3: Stub -->
    <div class="card">
      <div class="card-header">
        <div class="step-num">3</div>
        <h2>Stub (câu hỏi trong bảng)</h2>
        <span class="sub" id="stubCountLabel"></span>
      </div>
      <div class="card-body">
        <div class="stub-mode">
          <div class="mode-btn active" id="modeAll" onclick="setStubMode('all')">
            <div class="mb-icon">⚡</div>
            <div class="mb-name">Tất cả câu</div>
            <div class="mb-desc">Lấy toàn bộ câu codeable</div>
          </div>
          <div class="mode-btn" id="modeSelect" onclick="setStubMode('select')">
            <div class="mb-icon">☑️</div>
            <div class="mb-name">Chọn cụ thể</div>
            <div class="mb-desc">Tick từng câu muốn đưa vào</div>
          </div>
        </div>
        <div id="stubSelectArea" style="display:none">
          <input type="text" class="q-search" placeholder="🔍 Tìm câu hỏi..." oninput="filterList('stub', this.value)">
          <div class="q-list" id="stubList">
            <div class="q-select-all" onclick="toggleAll('stub')">
              <input type="checkbox" id="stubAllChk" onclick="event.stopPropagation()"> Chọn tất cả
            </div>
            <!-- populated by JS -->
          </div>
        </div>
        <div class="summary-bar" id="stubSummary"></div>
      </div>
    </div>

    <!-- Submit -->
    <div class="submit-section">
      <button class="btn-submit" onclick="submitForm()">✅ Xác nhận cấu hình</button>
      <div class="required-hint" id="requiredHint">⚠ Vui lòng chọn đủ 3 phần trên trước khi xác nhận</div>
    </div>

  </div><!-- /formArea -->

  <!-- Success screen -->
  <div class="success" id="successArea">
    <div class="s-icon">✅</div>
    <h2>Đã xác nhận!</h2>
    <p>Claude đã nhận được cấu hình.<br>Bạn có thể đóng panel này, Claude sẽ tạo datatable.json ngay.</p>
    <div class="s-summary" id="successSummary"></div>
  </div>

</div><!-- /page -->

<script>
// ── Data from Python ────────────────────────────────────────────────────────
const SURVEY_NAME  = {name_js};
const BANNER_QS    = {banner_js};
const STUB_QS      = {stub_js};

// ── State ───────────────────────────────────────────────────────────────────
window.__result   = null;
let stubMode      = 'all';
let bannerChecked = {{}};
let stubChecked   = {{}};

// ── Init ────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {{
  if (SURVEY_NAME) document.getElementById('surveyNameLabel').textContent = SURVEY_NAME + ' — Step 4';
  document.getElementById('stubCountLabel').textContent = STUB_QS.length + ' câu codeable';
  renderList('banner', BANNER_QS);
  renderList('stub',   STUB_QS);
}});

// ── Render question list ────────────────────────────────────────────────────
function renderList(which, qs, filter='') {{
  const listId  = which === 'banner' ? 'bannerList' : 'stubList';
  const checked = which === 'banner' ? bannerChecked : stubChecked;
  const container = document.getElementById(listId);

  // Keep the "select all" row
  const selectAllRow = container.querySelector('.q-select-all');
  // Remove old items
  [...container.querySelectorAll('.q-item')].forEach(el => el.remove());

  const fl = filter.toLowerCase();
  const visible = qs.filter(q =>
    !fl || q.label.toLowerCase().includes(fl) || q.title.toLowerCase().includes(fl)
  );

  visible.forEach(q => {{
    const isChecked = !!checked[q.label];
    const div = document.createElement('div');
    div.className = 'q-item' + (isChecked ? ' checked' : '');
    div.dataset.q = q.label;
    div.onclick = () => toggleItem(which, q.label, div);
    div.innerHTML = `
      <input type="checkbox" ${{isChecked ? 'checked' : ''}} onclick="event.stopPropagation();toggleItem('${{which}}','${{q.label}}',this.closest('.q-item'))">
      <span class="q-label">${{q.label}}</span>
      <span class="q-title">${{escHtml(q.title)}}</span>
      <span class="q-type">${{q.type}}</span>
    `;
    container.appendChild(div);
  }});

  updateSummary(which);
  updateAllChk(which, qs);
}}

function toggleItem(which, label, row) {{
  const checked = which === 'banner' ? bannerChecked : stubChecked;
  checked[label] = !checked[label];
  row.classList.toggle('checked', !!checked[label]);
  row.querySelector('input[type=checkbox]').checked = !!checked[label];
  updateSummary(which);
  const qs = which === 'banner' ? BANNER_QS : STUB_QS;
  updateAllChk(which, qs);
}}

function toggleAll(which) {{
  const qs      = which === 'banner' ? BANNER_QS : STUB_QS;
  const checked = which === 'banner' ? bannerChecked : stubChecked;
  const allChk  = document.getElementById(which === 'banner' ? 'bannerAllChk' : 'stubAllChk');
  const toCheck = !allChk.checked;
  qs.forEach(q => checked[q.label] = toCheck);
  allChk.checked = toCheck;
  renderList(which, qs);
}}

function updateAllChk(which, qs) {{
  const checked = which === 'banner' ? bannerChecked : stubChecked;
  const allChk  = document.getElementById(which === 'banner' ? 'bannerAllChk' : 'stubAllChk');
  if (!allChk) return;
  const total  = qs.length;
  const n      = qs.filter(q => checked[q.label]).length;
  allChk.checked       = total > 0 && n === total;
  allChk.indeterminate = n > 0 && n < total;
}}

function filterList(which, val) {{
  const qs = which === 'banner' ? BANNER_QS : STUB_QS;
  renderList(which, qs, val);
}}

function updateSummary(which) {{
  const checked = which === 'banner' ? bannerChecked : stubChecked;
  const sumEl   = document.getElementById(which + 'Summary');
  const labels  = Object.keys(checked).filter(k => checked[k]);
  if (labels.length === 0) {{
    sumEl.className = 'summary-bar';
    return;
  }}
  sumEl.className = 'summary-bar visible';
  const qs = (which === 'banner' ? BANNER_QS : STUB_QS);
  const parts = labels.map(l => {{
    const q = qs.find(x => x.label === l);
    return `<strong>${{l}}</strong> ${{q ? q.title : ''}}`;
  }});
  const prefix = which === 'banner'
    ? `<strong>Total</strong> (mặc định) + `
    : '';
  sumEl.innerHTML = `✓ Đã chọn ${{labels.length}} câu: ${{prefix}}${{parts.join(', ')}}`;
}}

// ── Table type ──────────────────────────────────────────────────────────────
function onTypeChange() {{
  document.querySelectorAll('.type-option').forEach((el, i) => {{
    const radio = el.querySelector('input[type=radio]');
    el.classList.toggle('selected', radio.checked);
  }});
}}

// ── Stub mode ───────────────────────────────────────────────────────────────
function setStubMode(mode) {{
  stubMode = mode;
  document.getElementById('modeAll').classList.toggle('active', mode === 'all');
  document.getElementById('modeSelect').classList.toggle('active', mode === 'select');
  document.getElementById('stubSelectArea').style.display = mode === 'select' ? 'block' : 'none';
  updateSummary('stub');
}}

// ── Submit ──────────────────────────────────────────────────────────────────
function submitForm() {{
  // Validate
  const typeVal = document.querySelector('input[name=table_type]:checked');
  if (!typeVal) {{
    document.getElementById('requiredHint').classList.add('show');
    return;
  }}
  document.getElementById('requiredHint').classList.remove('show');

  const bannerSelected = Object.keys(bannerChecked).filter(k => bannerChecked[k]);
  const stubSelected   = stubMode === 'all'
    ? STUB_QS.map(q => q.label)
    : Object.keys(stubChecked).filter(k => stubChecked[k]);

  const TYPE_LABELS = ['','Count only','Percentage only','Percentage + Sig test','Tất cả (Count + Pct + Sig)'];
  const result = {{
    submitted:  true,
    table_type: parseInt(typeVal.value),
    banner:     bannerSelected,
    stub_mode:  stubMode,
    stub:       stubSelected,
  }};

  window.__result = result;

  // Show success
  document.getElementById('formArea').style.display    = 'none';
  document.getElementById('successArea').className     = 'success show';
  document.getElementById('successSummary').innerHTML  = `
    <b>Loại bảng:</b> ${{TYPE_LABELS[result.table_type]}}<br>
    <b>Banner:</b> Total${{bannerSelected.length ? ' + ' + bannerSelected.join(', ') : ' (chỉ Total)'}}<br>
    <b>Stub:</b> ${{stubMode === 'all' ? 'Tất cả ' + STUB_QS.length + ' câu' : stubSelected.join(', ') || '(chưa chọn)'}}
  `;
}}

// ── Util ────────────────────────────────────────────────────────────────────
function escHtml(s) {{
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}}
</script>
</body>
</html>
"""

    out = pathlib.Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, encoding="utf-8")
    print(out.resolve())


# ─── CLI ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python tools/generate_step4_form.py <metadata.json> [output.html] [survey_name]")
        sys.exit(1)

    meta_path   = sys.argv[1]
    survey_name = sys.argv[3] if len(sys.argv) > 3 else pathlib.Path(meta_path).parent.parent.name

    if len(sys.argv) > 2:
        out_path = sys.argv[2]
    else:
        out_path = str(pathlib.Path(meta_path).parent.parent / "step4_form.html")

    generate_form(meta_path, out_path, survey_name)
