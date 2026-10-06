import io, os
from flask import Flask, request, jsonify, send_file, make_response
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

VERSION = "1.0"
API_TOKEN = os.environ.get("API_TOKEN", "").strip()
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

try:
    import pytesseract
    from PIL import Image
    LANGS = [l for l in pytesseract.get_languages(config="") if l != "osd"]
    OCR_OK = True
except Exception:
    LANGS, OCR_OK = [], False


@app.after_request
def cors(r):
    r.headers["Access-Control-Allow-Origin"] = "*"
    r.headers["Access-Control-Allow-Headers"] = "X-Token, Content-Type"
    r.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    return r


@app.route("/<path:p>", methods=["OPTIONS"])
@app.route("/", methods=["OPTIONS"])
def opts(p=""):
    return make_response("", 204)


def authed():
    return bool(API_TOKEN) and request.headers.get("X-Token", "").strip() == API_TOKEN


@app.route("/")
def home():
    return "STFFE server running"


@app.route("/health")
def health():
    return jsonify(ok=True, auth=authed(), version=VERSION, ocr=OCR_OK, langs=LANGS)


@app.route("/ocr", methods=["POST"])
def ocr():
    if not authed():
        return jsonify(ok=False, error="unauthorized"), 401
    if not OCR_OK:
        return jsonify(ok=False, error="tesseract missing"), 500
    f = request.files.get("file")
    if not f:
        return jsonify(ok=False, error="no file"), 400
    try:
        img = Image.open(f.stream).convert("RGB")
        lang = "+".join([l for l in ("eng", "tam") if l in LANGS]) or "eng"
        psm = request.args.get("psm", "").strip()
        cfg = "--oem 1" + (" --psm " + psm if psm.isdigit() else "")
        text = pytesseract.image_to_string(img, lang=lang, config=cfg)
        d = pytesseract.image_to_data(img, lang=lang, config=cfg,
                                      output_type=pytesseract.Output.DICT)
        c = [float(x) for x, t in zip(d["conf"], d["text"])
             if str(x) not in ("-1", "-1.0") and t.strip()]
        conf = sum(c) / len(c) if c else 0
        return jsonify(ok=True, text=text, confidence=conf)
    except Exception as e:
        return jsonify(ok=False, error=str(e)), 500


@app.route("/report/xlsx", methods=["POST"])
def xlsx():
    if not authed():
        return jsonify(ok=False, error="unauthorized"), 401
    spec = request.get_json(force=True, silent=True) or {}
    wb = Workbook()
    wb.remove(wb.active)
    for sh in spec.get("sheets", []) or [{"name": "Sheet1", "columns": [], "rows": []}]:
        ws = wb.create_sheet(str(sh.get("name") or "Sheet")[:31])
        cols = sh.get("columns", [])
        ws.append([c.get("header", "") for c in cols])
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in sh.get("rows", []):
            ws.append([row.get(c.get("key"), "") for c in cols])
        for i, c in enumerate(cols, 1):
            w = max([len(str(c.get("header", "")))] +
                    [len(str(r.get(c.get("key"), ""))) for r in sh.get("rows", [])])
            ws.column_dimensions[get_column_letter(i)].width = min(max(w + 2, 10), 45)
        for r in ws.iter_rows(min_row=2):
            for cell in r:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        ws.freeze_panes = "A2"
        if sh.get("landscape"):
            ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=spec.get("filename") or "report.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
