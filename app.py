import os, io, base64, json, re
from flask import Flask, request, send_file, jsonify, make_response
from flask_cors import CORS
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

app = Flask(__name__)
CORS(app)

DARK  = "1A1A2E"
GREEN = "1D9E75"
WHITE = "FFFFFF"
LGREY = "F7F7F7"

def hex_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def thin_border():
    s = Side(style="thin", color="DDDDDD")
    return Border(left=s, right=s, top=s, bottom=s)

def base64_to_pil(b64_str):
    if b64_str.startswith("data:"):
        b64_str = b64_str.split(",", 1)[1]
    raw = base64.b64decode(b64_str)
    return PILImage.open(io.BytesIO(raw)).convert("RGBA")

def pil_to_png_bytes(pil_img, size=(80, 80)):
    pil_img.thumbnail(size, PILImage.LANCZOS)
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    buf.seek(0)
    return buf

@app.after_request
def after_request(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type, Authorization'
    return response

@app.route('/')
def index():
    return jsonify({"status": "BGlam Excel API", "version": "1.1"})

@app.route('/excel', methods=['OPTIONS'])
@app.route('/extract-images', methods=['OPTIONS'])
def handle_options():
    return make_response('', 204)

# ── EXTRACT IMAGES FROM XLSX ──────────────────────────────────────────────────
@app.route('/extract-images', methods=['POST'])
def extract_images():
    """
    Receives a .xlsx file (multipart/form-data, field 'file').
    Returns JSON: { "BA-398": "data:image/png;base64,...", ... }
    """
    if 'file' not in request.files:
        return jsonify({"error": "No file provided"}), 400

    f = request.files['file']
    try:
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(f.read()))
        ws = wb.active

        # Auto-detect header row and ref column
        header_row = 1
        ref_col = 1  # default col A
        REF_KEYWORDS = ('item no', 'item no.', 'item no:', 'ref', 'reference', 'sku', 'code')
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=15, values_only=False), 1):
            for cell in row:
                if cell.value and str(cell.value).strip().lower() in REF_KEYWORDS:
                    header_row = i
                    ref_col = cell.column
                    break
            else:
                continue
            break

        # Build row -> ref map using detected column
        row_to_ref = {}
        for row in ws.iter_rows(min_row=header_row+1, values_only=False):
            for cell in row:
                if cell.column == ref_col and cell.value:
                    val = str(cell.value).strip()
                    if val.lower() not in ('item no', 'item no.', 'item no:', 'ref', 'reference', ''):
                        row_to_ref[cell.row] = val
                        break

        # Build image row -> base64 map
        img_map = {}
        for img in ws._images:
            try:
                anchor = img.anchor
                if hasattr(anchor, '_from'):
                    img_row = anchor._from.row + 1
                    data = img._data()
                    if data and len(data) > 100:
                        b64 = 'data:image/png;base64,' + base64.b64encode(data).decode('utf-8')
                        img_map[img_row] = b64
            except Exception:
                pass

        # Match images to refs (search ±5 rows around image anchor)
        result = {}
        for img_row, b64 in img_map.items():
            ref = row_to_ref.get(img_row)
            if not ref:
                for offset in range(-5, 6):
                    ref = row_to_ref.get(img_row + offset)
                    if ref:
                        break
            if ref and ref not in result:
                result[ref] = b64

        return jsonify({
            "images": result,
            "matched": len(result),
            "total_images": len(img_map),
            "total_refs": len(row_to_ref)
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── GENERATE EXCEL WITH IMAGES ────────────────────────────────────────────────
@app.route('/excel', methods=['POST'])
def generate_excel():
    try:
        data = request.get_json(force=True)
    except Exception:
        return jsonify({"error": "Invalid JSON"}), 400

    supplier   = data.get("supplier",   "BGlam")
    order_name = data.get("order_name", "Commande")
    date_str   = data.get("date",       "")
    items      = data.get("items",      [])

    if not items:
        return jsonify({"error": "No items provided"}), 400

    wb = Workbook()
    ws = wb.active
    ws.title = "Commande"

    # Title row
    ws.merge_cells("A1:K1")
    ws["A1"] = f"{supplier} — Bon de Commande"
    ws["A1"].font      = Font(name="Arial", size=14, bold=True, color=WHITE)
    ws["A1"].fill      = hex_fill(DARK)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # Meta row
    ws.merge_cells("A2:D2")
    ws["A2"] = f"Commande : {order_name}"
    ws["A2"].font = Font(name="Arial", size=11, bold=True)
    ws["A2"].fill = hex_fill(LGREY)
    ws["A2"].alignment = Alignment(vertical="center")
    ws.merge_cells("E2:H2")
    ws["E2"] = f"Date : {date_str}"
    ws["E2"].font = Font(name="Arial", size=11)
    ws["E2"].fill = hex_fill(LGREY)
    ws["E2"].alignment = Alignment(vertical="center")
    ws.row_dimensions[2].height = 22

    # Header row
    headers    = ["Photo","CTN NO","ITEM NO","DESCRIPTION","CATEGORIE","PRICE","PCS/CTN","CTN","QTY","AMOUNT","REMARK"]
    col_widths = [14,      10,      12,       28,            22,         10,     10,       6,    8,    12,      35]
    HEADER_ROW = 3
    for col_idx, (h, w) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=HEADER_ROW, column=col_idx, value=h)
        cell.font      = Font(name="Arial", size=10, bold=True, color=WHITE)
        cell.fill      = hex_fill(DARK)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = thin_border()
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.row_dimensions[HEADER_ROW].height = 20

    IMG_ROW_HEIGHT = 62
    IMG_SIZE       = (60, 60)
    total_amount   = 0.0

    for row_offset, item in enumerate(items):
        r     = HEADER_ROW + 1 + row_offset
        qty   = item.get("qty") or 0
        price = item.get("price")
        amt   = round(qty * price, 2) if price is not None else None
        if amt:
            total_amount += amt
        bg = LGREY if row_offset % 2 == 0 else WHITE

        values = [
            item.get("group",  ""),
            item.get("ref",    ""),
            item.get("desc",   ""),
            item.get("cat",    ""),
            price,
            item.get("moq"),
            "",
            qty,
            amt,
            item.get("remark", ""),
        ]
        for col_idx, val in enumerate(values, start=2):
            cell = ws.cell(row=r, column=col_idx, value=val)
            cell.font      = Font(name="Arial", size=10)
            cell.fill      = hex_fill(bg)
            cell.alignment = Alignment(vertical="center")
            cell.border    = thin_border()
            if col_idx in (6, 7, 9, 10):
                cell.alignment = Alignment(horizontal="right", vertical="center")
            if col_idx == 10 and amt is not None:
                cell.font = Font(name="Arial", size=10, bold=True, color=GREEN)

        ws.row_dimensions[r].height = IMG_ROW_HEIGHT

        # Embed image in col A
        img_b64 = item.get("img", "")
        if img_b64 and len(img_b64) > 100:
            try:
                pil_img = base64_to_pil(img_b64)
                png_buf = pil_to_png_bytes(pil_img, size=IMG_SIZE)
                xl_img  = XLImage(png_buf)
                xl_img.anchor = f"A{r}"
                ws.add_image(xl_img)
            except Exception:
                pass
        cell = ws.cell(row=r, column=1, value="")
        cell.fill   = hex_fill(bg)
        cell.border = thin_border()

    # Total row
    total_row = HEADER_ROW + 1 + len(items)
    total_qty = sum((item.get("qty") or 0) for item in items)
    ws.merge_cells(f"A{total_row}:H{total_row}")
    cell = ws.cell(row=total_row, column=1, value="TOTAL")
    cell.font      = Font(name="Arial", size=11, bold=True, color=WHITE)
    cell.fill      = hex_fill(DARK)
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell = ws.cell(row=total_row, column=9, value=total_qty)
    cell.font      = Font(name="Arial", size=11, bold=True, color=WHITE)
    cell.fill      = hex_fill(DARK)
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell = ws.cell(row=total_row, column=10, value=round(total_amount, 2))
    cell.font      = Font(name="Arial", size=12, bold=True, color="4ADE80")
    cell.fill      = hex_fill(DARK)
    cell.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[total_row].height = 24
    ws.freeze_panes = f"B{HEADER_ROW + 1}"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    safe_name = re.sub(r"[^a-zA-Z0-9_\- ]", "_", f"{supplier}_{order_name}")
    filename  = f"{safe_name}_{date_str.replace('/', '-')}.xlsx"

    response = make_response(send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ))
    response.headers['Access-Control-Allow-Origin'] = '*'
    return response

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
