"""Web page for converting PDFs: upload a PDF, download an Excel workbook.

Run with:  python -m pdf2excel.web   then open http://127.0.0.1:5000
"""

from __future__ import annotations

import io
import tempfile
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from werkzeug.utils import secure_filename

from .converter import convert

MAX_UPLOAD_MB = 50

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
          --accent:#1f7a4d; --accent-soft:#e7f4ed; --error:#b42318; }
  @media (prefers-color-scheme: dark) {
    :root { --bg:#14171c; --card:#1d2128; --text:#e6e8ec; --muted:#98a2b3; --line:#363c47;
            --accent:#3ccf8e; --accent-soft:#17352a; --error:#f97066; }
  }
  * { box-sizing: border-box; }
  body { margin:0; font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
         background:var(--bg); color:var(--text); }
  main { max-width:560px; margin:48px auto; padding:0 16px; }
  h1 { font-size:1.6rem; margin:0 0 4px; }
  p.sub { color:var(--muted); margin:0 0 24px; }
  .card { background:var(--card); border:1px solid var(--line); border-radius:12px; padding:20px; }
  #drop { border:2px dashed var(--line); border-radius:10px; padding:36px 16px; text-align:center;
          cursor:pointer; transition: background .15s, border-color .15s; }
  #drop.over { border-color:var(--accent); background:var(--accent-soft); }
  #drop strong { display:block; font-size:1.05rem; margin-bottom:4px; }
  #drop span { color:var(--muted); font-size:.9rem; }
  #file { display:none; }
  .opts { display:grid; gap:12px; margin-top:18px; }
  label { font-size:.92rem; }
  input[type=text] { width:100%; padding:8px 10px; border:1px solid var(--line); border-radius:8px;
                     background:var(--bg); color:var(--text); font:inherit; margin-top:4px; }
  button { margin-top:18px; width:100%; padding:12px; border:0; border-radius:8px;
           background:var(--accent); color:#fff; font:inherit; font-weight:600; cursor:pointer; }
  button:disabled { opacity:.5; cursor:default; }
  #status { margin-top:14px; min-height:1.4em; font-size:.92rem; }
  #status.error { color:var(--error); }
  #status.ok { color:var(--accent); }
</style>
</head>
<body>
<main>
  <h1>PDF to Excel</h1>
  <p class="sub">Upload a PDF and get its tables back as an Excel workbook.</p>
  <div class="card">
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
    <button id="go" disabled>Convert to Excel</button>
    <div id="status" role="status"></div>
  </div>
</main>
<script>
const drop = document.getElementById('drop'), input = document.getElementById('file'),
      go = document.getElementById('go'), statusEl = document.getElementById('status');
let chosen = null;

function pick(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith('.pdf')) { setStatus('Please choose a .pdf file.', 'error'); return; }
  chosen = file;
  document.getElementById('dropTitle').textContent = file.name;
  document.getElementById('dropHint').textContent = (file.size < 1048576 ? Math.max(1, Math.round(file.size / 1024)) + ' KB' : (file.size / 1048576).toFixed(1) + ' MB') + '. Click to choose a different file.';
  go.disabled = false;
  setStatus('');
}
function setStatus(msg, kind) { statusEl.textContent = msg; statusEl.className = kind || ''; }

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
  form.append('pages', document.getElementById('pages').value);
  if (document.getElementById('single').checked) form.append('single_sheet', '1');
  if (document.getElementById('keepText').checked) form.append('keep_text', '1');
  go.disabled = true;
  setStatus('Converting...');
  try {
    const res = await fetch('/convert', { method: 'POST', body: form });
    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: 'Conversion failed.' }));
      throw new Error(err.error);
    }
    const blob = await res.blob();
    const name = chosen.name.replace(/\\.pdf$/i, '') + '.xlsx';
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    const count = res.headers.get('X-Table-Count');
    setStatus('Done: ' + name + ' downloaded (' + count + (count === '1' ? ' table).' : ' tables).'), 'ok');
  } catch (e) {
    setStatus(e.message, 'error');
  } finally {
    go.disabled = false;
  }
});
</script>
</body>
</html>"""


@app.get("/")
def index():
    return PAGE


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

    response = send_file(
        data,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=pdf_path.stem + ".xlsx",
    )
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
