"""Web page: paste text or upload a PDF, download an Excel workbook.

Run with:  python -m pdf2excel.web   then open http://127.0.0.1:5000
"""

from __future__ import annotations

import io
import tempfile
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from werkzeug.utils import secure_filename

from .converter import convert, rows_to_workbook
from .text_parser import DELIMITERS, parse_text

MAX_UPLOAD_MB = 50
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

PAGE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PDF to Excel</title>
<style>
  :root { --bg:#f6f7f9; --card:#fff; --text:#1d2330; --muted:#667085; --line:#d0d5dd;
          --accent:#1f7a4d; --accent-soft:#e7f4ed; --error:#b42318; --head:#eef4fb; }
  @media (prefers-color-scheme: dark) {
    :root { --bg:#14171c; --card:#1d2128; --text:#e6e8ec; --muted:#98a2b3; --line:#363c47;
            --accent:#3ccf8e; --accent-soft:#17352a; --error:#f97066; --head:#232a35; }
  }
  * { box-sizing: border-box; }
  body { margin:0; font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
         background:var(--bg); color:var(--text); }
  main { max-width:860px; margin:40px auto; padding:0 16px; }
  h1 { font-size:1.6rem; margin:0 0 4px; }
  p.sub { color:var(--muted); margin:0 0 20px; }
  .tabs { display:flex; gap:4px; margin-bottom:-1px; }
  .tab { padding:10px 16px; border:1px solid transparent; border-bottom:0; border-radius:10px 10px 0 0;
         background:none; color:var(--muted); font:inherit; font-weight:600; cursor:pointer; margin:0; width:auto; }
  .tab[aria-selected=true] { background:var(--card); border-color:var(--line); color:var(--text); }
  .card { background:var(--card); border:1px solid var(--line); border-radius:0 12px 12px 12px; padding:20px; }
  textarea { width:100%; min-height:200px; padding:12px; border:1px solid var(--line); border-radius:8px;
             background:var(--bg); color:var(--text); font:13px/1.45 ui-monospace, Consolas, monospace; resize:vertical; }
  .row { display:flex; flex-wrap:wrap; gap:12px 20px; align-items:center; margin-top:14px; }
  select, input[type=text] { padding:7px 10px; border:1px solid var(--line); border-radius:8px;
                             background:var(--bg); color:var(--text); font:inherit; }
  label { font-size:.92rem; }
  #drop { border:2px dashed var(--line); border-radius:10px; padding:36px 16px; text-align:center;
          cursor:pointer; transition: background .15s, border-color .15s; }
  #drop.over { border-color:var(--accent); background:var(--accent-soft); }
  #drop strong { display:block; font-size:1.05rem; margin-bottom:4px; }
  #drop span { color:var(--muted); font-size:.9rem; }
  #file { display:none; }
  .opts { display:grid; gap:12px; margin-top:18px; }
  .opts input[type=text] { width:100%; margin-top:4px; }
  button.go { margin-top:18px; width:100%; padding:12px; border:0; border-radius:8px;
              background:var(--accent); color:#fff; font:inherit; font-weight:600; cursor:pointer; }
  button.go:disabled { opacity:.5; cursor:default; }
  .actions { display:flex; flex-wrap:wrap; gap:10px; margin-top:18px; }
  .actions button.go { margin-top:0; flex:1 1 220px; width:auto; }
  button.secondary { padding:12px 16px; border:1px solid var(--line); border-radius:8px; background:var(--card);
                     color:var(--text); font:inherit; font-weight:600; cursor:pointer; flex:0 1 auto; }
  button.secondary:hover { border-color:var(--accent); }
  .status { margin-top:14px; min-height:1.4em; font-size:.92rem; }
  .status.error { color:var(--error); }
  .status.ok { color:var(--accent); }
  .previewHead { display:flex; justify-content:space-between; align-items:baseline; margin-top:18px; }
  .previewHead h2 { font-size:1rem; margin:0; }
  #summary { color:var(--muted); font-size:.88rem; }
  .preview { margin-top:8px; max-height:340px; overflow:auto; border:1px solid var(--line); border-radius:8px; }
  .preview:empty { display:none; }
  table { border-collapse:collapse; font-size:.88rem; width:max-content; min-width:100%; }
  td, th { border-bottom:1px solid var(--line); border-right:1px solid var(--line); padding:6px 10px;
           text-align:left; white-space:nowrap; max-width:320px; overflow:hidden; text-overflow:ellipsis; }
  th { background:var(--head); position:sticky; top:0; }
  td.num { text-align:right; font-variant-numeric: tabular-nums; }
  .empty { color:var(--muted); font-size:.9rem; padding:16px 0 0; }
  [hidden] { display:none !important; }
</style>
</head>
<body>
<main>
  <h1>PDF to Excel</h1>
  <p class="sub">Paste text or upload a PDF and get a tidy Excel spreadsheet back.</p>
  <div class="tabs" role="tablist">
    <button class="tab" role="tab" id="tab-text" aria-selected="true">Paste text</button>
    <button class="tab" role="tab" id="tab-pdf" aria-selected="false">Upload PDF</button>
  </div>

  <section class="card" id="panel-text">
    <textarea id="text" placeholder="Paste anything with rows and columns here: a table copied from a PDF, a web page, an email, a CSV..." spellcheck="false"></textarea>
    <div class="row">
      <label>Split columns by
        <select id="delim">
          <option value="auto">Detect automatically</option>
          <option value="tab">Tabs</option>
          <option value="comma">Commas</option>
          <option value="semicolon">Semicolons</option>
          <option value="pipe">| Pipes</option>
          <option value="spaces">2+ spaces</option>
          <option value="numbers">Text, then numbers at the end</option>
          <option value="label">Label: value</option>
        </select>
      </label>
      <label><input id="hasHeader" type="checkbox" checked> First row is a header</label>
      <label><input id="textKeep" type="checkbox"> Keep numbers as text</label>
    </div>
    <div class="previewHead"><h2>Preview</h2><span id="summary"></span></div>
    <div class="empty" id="emptyMsg">Your spreadsheet will appear here as you paste.</div>
    <div class="preview" id="preview"></div>
    <div class="actions">
      <button class="go" id="textGo" disabled>Download Excel</button>
      <button class="secondary" id="pasteNew" title="Replace everything with what you've copied">Paste new data</button>
      <button class="secondary" id="startOver" title="Clear the text and preview">Start over</button>
    </div>
    <div class="status" id="textStatus" role="status"></div>
  </section>

  <section class="card" id="panel-pdf" hidden>
    <div id="drop" tabindex="0">
      <strong id="dropTitle">Drop a PDF here or click to choose</strong>
      <span id="dropHint">Up to """ + str(MAX_UPLOAD_MB) + """ MB. PDFs with selectable text work best.</span>
    </div>
    <input id="file" type="file" accept="application/pdf,.pdf">
    <div class="opts">
      <label>Pages (optional, e.g. 1,3-5)
        <input id="pages" type="text" placeholder="All pages">
      </label>
      <label><input id="single" type="checkbox"> Put all tables on one sheet</label>
      <label><input id="keepText" type="checkbox"> Keep numbers as text</label>
    </div>
    <button class="go" id="go" disabled>Convert to Excel</button>
    <div class="status" id="status" role="status"></div>
  </section>
</main>
<script>
const $ = id => document.getElementById(id);
function setStatus(el, msg, kind) { el.textContent = msg; el.className = 'status ' + (kind || ''); }
function download(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}
async function errorOf(res) {
  const err = await res.json().catch(() => ({}));
  return new Error(err.error || 'Something went wrong.');
}

// ---- tabs ----
for (const [tab, panel] of [['tab-text','panel-text'], ['tab-pdf','panel-pdf']]) {
  $(tab).addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(t => t.setAttribute('aria-selected', t.id === tab));
    $('panel-text').hidden = panel !== 'panel-text';
    $('panel-pdf').hidden = panel !== 'panel-pdf';
  });
}

// ---- paste text ----
const MODE_NAMES = { tab:'tabs', comma:'commas', semicolon:'semicolons', pipe:'pipes',
  spaces:'2+ spaces', numbers:'text, then numbers', label:'label: value' };
let parsed = null, timer = null, userTouchedHeader = false;

function textBody() {
  return JSON.stringify({ text: $('text').value, delimiter: $('delim').value,
                          keep_text: $('textKeep').checked, header: $('hasHeader').checked });
}
function renderPreview(data) {
  const box = $('preview');
  box.innerHTML = '';
  const rows = data.rows;
  if (!rows.length) { $('emptyMsg').hidden = false; $('summary').textContent = ''; $('textGo').disabled = true; return; }
  $('emptyMsg').hidden = true;
  const table = document.createElement('table');
  const header = $('hasHeader').checked;
  rows.slice(0, 200).forEach((row, i) => {
    const tr = document.createElement('tr');
    row.forEach(cell => {
      const td = document.createElement(header && i === 0 ? 'th' : 'td');
      td.textContent = cell === null ? '' : cell;
      td.title = td.textContent;
      if (typeof cell === 'number') td.className = 'num';
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
  box.appendChild(table);
  const cols = rows[0].length;
  $('summary').textContent = rows.length + (rows.length === 1 ? ' row' : ' rows') + ', ' + cols +
    (cols === 1 ? ' column' : ' columns') + ' · split by ' + MODE_NAMES[data.mode] +
    (rows.length > 200 ? ' · showing first 200 rows' : '');
  $('textGo').disabled = false;
}
async function refresh() {
  if (!$('text').value.trim()) { renderPreview({ rows: [] }); return; }
  try {
    const res = await fetch('/parse-text', { method:'POST', headers:{'Content-Type':'application/json'}, body: textBody() });
    if (!res.ok) throw await errorOf(res);
    parsed = await res.json();
    if (!userTouchedHeader) $('hasHeader').checked = parsed.header;
    renderPreview(parsed);
    setStatus($('textStatus'), '');
  } catch (e) { setStatus($('textStatus'), e.message, 'error'); }
}
function schedule() { clearTimeout(timer); timer = setTimeout(refresh, 250); }
$('text').addEventListener('input', () => { userTouchedHeader = false; schedule(); });
$('delim').addEventListener('change', refresh);
$('textKeep').addEventListener('change', refresh);
$('hasHeader').addEventListener('change', () => { userTouchedHeader = true; if (parsed) renderPreview(parsed); });

function resetText(newText) {
  clearTimeout(timer);
  parsed = null;
  userTouchedHeader = false;
  $('delim').value = 'auto';
  $('textKeep').checked = false;
  $('hasHeader').checked = true;
  $('text').value = newText || '';
  setStatus($('textStatus'), '');
  renderPreview({ rows: [] });
  $('preview').scrollTop = 0;
  if (newText) refresh();
  $('text').focus();
}
$('startOver').addEventListener('click', () => resetText(''));
$('pasteNew').addEventListener('click', async () => {
  try {
    const clip = await navigator.clipboard.readText();
    if (!clip.trim()) throw new Error('empty');
    resetText(clip);
  } catch (e) {
    // clipboard blocked or empty: clear the box so the user can paste with Ctrl+V
    resetText('');
    setStatus($('textStatus'), e.message === 'empty'
      ? 'Your clipboard is empty. Copy some data, then try again.'
      : 'Cleared. Press Ctrl+V to paste your new data.');
  }
});
// pasting over a box that already holds data replaces it and starts fresh
$('text').addEventListener('paste', e => {
  const box = $('text');
  const all = box.selectionStart === 0 && box.selectionEnd === box.value.length;
  if (box.value.trim() && all) {
    e.preventDefault();
    resetText(e.clipboardData.getData('text'));
  }
});

$('textGo').addEventListener('click', async () => {
  $('textGo').disabled = true;
  setStatus($('textStatus'), 'Building spreadsheet...');
  try {
    const res = await fetch('/convert-text', { method:'POST', headers:{'Content-Type':'application/json'}, body: textBody() });
    if (!res.ok) throw await errorOf(res);
    download(await res.blob(), 'pasted-table.xlsx');
    setStatus($('textStatus'), 'Done: pasted-table.xlsx downloaded.', 'ok');
  } catch (e) { setStatus($('textStatus'), e.message, 'error'); }
  finally { $('textGo').disabled = false; }
});

// ---- upload PDF ----
const drop = $('drop'), input = $('file'), go = $('go');
let chosen = null;
function pick(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith('.pdf')) { setStatus($('status'), 'Please choose a .pdf file.', 'error'); return; }
  chosen = file;
  $('dropTitle').textContent = file.name;
  $('dropHint').textContent = (file.size < 1048576 ? Math.max(1, Math.round(file.size / 1024)) + ' KB'
    : (file.size / 1048576).toFixed(1) + ' MB') + '. Click to choose a different file.';
  go.disabled = false;
  setStatus($('status'), '');
}
drop.addEventListener('click', () => input.click());
drop.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') input.click(); });
input.addEventListener('change', () => pick(input.files[0]));
['dragenter','dragover'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave','drop'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => pick(e.dataTransfer.files[0]));

go.addEventListener('click', async () => {
  if (!chosen) return;
  const form = new FormData();
  form.append('pdf', chosen);
  form.append('pages', $('pages').value);
  if ($('single').checked) form.append('single_sheet', '1');
  if ($('keepText').checked) form.append('keep_text', '1');
  go.disabled = true;
  setStatus($('status'), 'Converting...');
  try {
    const res = await fetch('/convert', { method: 'POST', body: form });
    if (!res.ok) throw await errorOf(res);
    const name = chosen.name.replace(/\\.pdf$/i, '') + '.xlsx';
    download(await res.blob(), name);
    const count = res.headers.get('X-Table-Count');
    setStatus($('status'), 'Done: ' + name + ' downloaded (' + count + (count === '1' ? ' table).' : ' tables).'), 'ok');
  } catch (e) { setStatus($('status'), e.message, 'error'); }
  finally { go.disabled = false; }
});
</script>
</body>
</html>"""


@app.get("/")
def index():
    return PAGE


def _text_request() -> tuple[dict, dict] | tuple[None, tuple]:
    body = request.get_json(silent=True) or {}
    text = body.get("text") or ""
    delimiter = body.get("delimiter") or "auto"
    if delimiter not in DELIMITERS:
        return None, (jsonify(error=f"Unknown column separator '{delimiter}'."), 400)
    result = parse_text(text, delimiter, convert_numbers=not body.get("keep_text"))
    return body, result


@app.post("/parse-text")
def parse_text_preview():
    body, result = _text_request()
    if body is None:
        return result
    return jsonify(result)


@app.post("/convert-text")
def convert_text():
    body, result = _text_request()
    if body is None:
        return result
    if not result["rows"]:
        return jsonify(error="Paste some text first."), 400
    header = body.get("header", result["header"])
    wb = rows_to_workbook(result["rows"], header=bool(header), title="Pasted Table")
    data = io.BytesIO()
    wb.save(data)
    data.seek(0)
    return send_file(data, mimetype=XLSX_MIME, as_attachment=True, download_name="pasted-table.xlsx")


@app.post("/convert")
def convert_upload():
    upload = request.files.get("pdf")
    if upload is None or not upload.filename:
        return jsonify(error="No file was uploaded."), 400
    name = secure_filename(upload.filename) or "upload.pdf"
    if not name.lower().endswith(".pdf"):
        return jsonify(error="Please upload a .pdf file."), 400

    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = Path(tmp) / name
        upload.save(pdf_path)
        try:
            out_path, count = convert(
                pdf_path,
                Path(tmp) / (pdf_path.stem + ".xlsx"),
                pages=request.form.get("pages") or None,
                single_sheet=bool(request.form.get("single_sheet")),
                convert_numbers=not request.form.get("keep_text"),
            )
        except ValueError as exc:  # bad page range
            return jsonify(error=str(exc)), 400
        except Exception:
            return jsonify(error="Could not read this PDF. It may be damaged or password-protected."), 400
        data = io.BytesIO(out_path.read_bytes())

    response = send_file(data, mimetype=XLSX_MIME, as_attachment=True, download_name=pdf_path.stem + ".xlsx")
    response.headers["X-Table-Count"] = str(count)
    return response


@app.errorhandler(413)
def too_large(_):
    return jsonify(error=f"File is larger than {MAX_UPLOAD_MB} MB."), 413


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="Run the PDF to Excel web page.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=5000)
    args = parser.parse_args()
    print(f"Open http://{args.host}:{args.port} in your browser. Press Ctrl+C to stop.")
    app.run(host=args.host, port=args.port)


if __name__ == "__main__":
    main()
