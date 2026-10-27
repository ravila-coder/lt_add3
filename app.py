"""Loan tape CSV -> Portfolio Analysis.xlsx.

Run:  python app.py           (http://127.0.0.1:5000)
Check: python app.py --selftest
"""
import csv, io, os, re, sqlite3, sys, tempfile, uuid, zipfile
from copy import copy
from datetime import date, datetime

from flask import Flask, jsonify, render_template, request, send_file
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

HERE = os.path.dirname(os.path.abspath(__file__))
# the web build ships it as template.xlsx; the browser engine writes it back under the old name
TEMPLATE = next((p for p in (os.path.join(HERE, "Portfolio Analysis.xlsx"),
                             os.path.join(HERE, "template.xlsx")) if os.path.exists(p)),
                os.path.join(HERE, "Portfolio Analysis.xlsx"))
ACCT = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
MMMYY = "[$-409]mmm\\-yy;@"

app = Flask(__name__)

# ponytail: one global dataset, single-user local tool. Key by session id if it ever goes multi-user.
DB = None          # sqlite connection, in-memory
RAW = []           # original CSV rows as text, never mutated
HEADERS = []       # cleaned column names, in CSV order
TYPES = {}         # name -> one of string/integer/float/date/boolean
CALCS = []         # [(name, sql_expr)] applied after every rebuild
PAY = {}           # 'esperados'/'recibidos' -> {headers, types, rows}
EMPTY_ROWS = {}    # slot -> fully empty rows dropped at ingest, for the audit tab
FILES = {}         # download id -> path


# ---------- parsing (Mexican formats) ----------

def parse_number(s):
    """'$ 1,234.56' / '1.234,56' / '(1,234.56)' / '45%' -> float, else None."""
    t = (s or "").strip().replace("\u00a0", " ")   # NBSP, common in Excel exports
    if not t:
        return None
    neg = t.startswith("(") and t.endswith(")")
    if neg:
        t = t[1:-1].strip()
    pct = t.endswith("%")
    if pct:
        t = t[:-1].strip()
    t = re.sub(r"(?i)\b(MXN|USD|MN|EUR|pesos?)\b", "", t)
    t = t.replace("$", "").replace(" ", "")
    # shape check first: without it '18/09/2026' strips down to a valid-looking integer
    if not re.fullmatch(r"[+-]?\d[\d.,]*", t):
        return None
    has_c, has_d = "," in t, "." in t
    if has_c and has_d:
        dec = "," if t.rfind(",") > t.rfind(".") else "."
    elif has_c:
        parts = t.split(",")
        # one comma with a group that is not 3 digits is a decimal: '1,5' -> 1.5, '1,234' -> 1234
        dec = "," if len(parts) == 2 and len(parts[1]) != 3 else None
    elif has_d:
        # In Mexico the dot IS the decimal separator, so a lone dot is always decimal.
        # Only repeated dots ('1.234.567') mean European thousands grouping. Treating
        # '0.425' as a grouped thousand would silently multiply the value by 1000.
        dec = "." if t.count(".") == 1 else None
    else:
        dec = None
    if dec:
        t = t.replace("," if dec == "." else ".", "").replace(dec, ".")
    else:
        t = t.replace(",", "").replace(".", "")
    try:
        v = float(t)
    except ValueError:
        return None
    if neg:
        v = -abs(v)
    return v / 100 if pct else v


DATE_FMTS = ("%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y",
             "%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y", "%Y%m%d")


def parse_date(s):
    t = (s or "").strip()
    if not t:
        return None
    t = t.split(" ")[0] if " " in t and ":" in t else t     # drop a time part
    for f in DATE_FMTS:
        try:
            return datetime.strptime(t, f).date()
        except ValueError:
            pass
    return None


TRUE = {"true", "verdadero", "si", "sí", "yes", "y", "1"}
FALSE = {"false", "falso", "no", "n", "0"}


def parse_bool(s):
    t = (s or "").strip().lower()
    return True if t in TRUE else False if t in FALSE else None


def parse_int(s):
    v = parse_number(s)
    return int(v) if v is not None and float(v).is_integer() else None


PARSERS = {"integer": parse_int, "float": parse_number,
           "date": parse_date, "boolean": parse_bool, "string": lambda s: s}


def detect_type(values):
    vals = [v for v in values if (v or "").strip()]
    if not vals:
        return "string"
    for t in ("integer", "float", "date", "boolean"):
        p = PARSERS[t]
        if all(p(v) is not None for v in vals):
            return t
    return "string"


def clean_name(n, taken):
    """Excel structured references choke on [ ] # ' @ - strip them, then dedupe."""
    n = re.sub(r"[\[\]#'\"@]", "", (n or "").strip()).strip() or "col"
    base, i = n, 2
    while n.lower() in taken:
        n, i = f"{base}_{i}", i + 1
    taken.add(n.lower())
    return n


# ---------- sqlite ----------

SQL_TYPE = {"integer": "INTEGER", "float": "REAL", "date": "TEXT",
            "boolean": "INTEGER", "string": "TEXT"}


def q(name):
    return '"' + name.replace('"', '') + '"'


def key_ok(col):
    """SQL predicate: col holds a usable credit id. A blank id is a missing id, not a
    key. Left in, a '' in a text-typed tape column would join to every blank-id
    payment, put those rows in the workbook and count them as both null and matched."""
    return f"{q(col)} IS NOT NULL AND TRIM(CAST({q(col)} AS TEXT)) <> ''"


def tape_ids(col):
    """The tape's distinct usable credit ids, as text, for an IN (...) join."""
    return f"(SELECT DISTINCT CAST({q(col)} AS TEXT) FROM tape WHERE {key_ok(col)})"


def rebuild():
    """Rebuild the typed `tape` table from RAW, then re-apply calculated columns.

    Conversions always start from the untouched CSV text, so a retype never loses data.
    Returns {column: failed_row_count} for the typed columns.
    """
    global DB
    # Flask serves each request on a new thread; sqlite connections are thread-bound by default.
    # ponytail: no lock, single-user local tool. Add one if this ever serves concurrent users.
    DB = sqlite3.connect(":memory:", check_same_thread=False)
    cols = ", ".join(f"{q(h)} {SQL_TYPE[TYPES[h]]}" for h in HEADERS)
    DB.execute(f"CREATE TABLE tape ({cols})")
    fails = {h: 0 for h in HEADERS}
    rows = []
    for raw in RAW:
        out = []
        for h, txt in zip(HEADERS, raw):
            t = TYPES[h]
            if t == "string":
                out.append(txt)
                continue
            if not (txt or "").strip():
                out.append(None)
                continue
            v = PARSERS[t](txt)
            if v is None:
                fails[h] += 1
                out.append(None)
            elif t == "date":
                out.append(v.isoformat())
            elif t == "boolean":
                out.append(1 if v else 0)
            else:
                out.append(v)
        rows.append(out)
    DB.executemany(
        f"INSERT INTO tape VALUES ({','.join('?' * len(HEADERS))})", rows)
    DB.commit()
    for name, expr in CALCS:
        apply_calc(name, expr)
    # the payment tables live in the same connection; a retype or a calculated column
    # after the payment uploads would otherwise silently drop them
    for slot, meta in list(PAY.items()):
        load_pay(slot, meta["headers"], meta["raw"])
    return fails


def apply_calc(name, expr):
    DB.execute(f"ALTER TABLE tape ADD COLUMN {q(name)}")
    DB.execute(f"UPDATE tape SET {q(name)} = ({expr})")
    DB.commit()


def calc_type(name):
    # scan every row, not a sample: a CASE WHEN that is mostly 0 would otherwise
    # look like an integer and display its rare decimal values rounded
    vals = [r[0] for r in
            DB.execute(f"SELECT {q(name)} FROM tape WHERE {q(name)} IS NOT NULL")]
    if not vals:
        return "string"
    if any(isinstance(v, float) for v in vals):
        return "float"
    if all(isinstance(v, int) for v in vals):
        return "integer"
    return detect_type([str(v) for v in vals])


def bubbles():
    return [{"name": h, "type": TYPES[h], "calc": False} for h in HEADERS] + \
           [{"name": n, "type": TYPES[n], "calc": True} for n, _ in CALCS]


# ---------- routes ----------

@app.get("/")
def index():
    return render_template("index.html")


def cell_text(v):
    """openpyxl hands back typed values; the type pipeline works on the raw text."""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.date().isoformat() if not (v.hour or v.minute) else v.isoformat(" ")
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def read_table(filename, data, slot="tape"):
    """(headers, rows-of-text) from a .csv/.txt or .xlsx upload.

    Rows with no value in any cell are dropped here and counted in EMPTY_ROWS, so the
    audit tab can say how many never reached the workbook."""
    if filename.lower().endswith((".xlsx", ".xlsm")):
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        it = wb[wb.sheetnames[0]].iter_rows(values_only=True)
        rows = [[cell_text(c) for c in r] for r in it]
        wb.close()
    else:
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                text = data.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError("Could not decode the file")
        try:
            dialect = csv.Sniffer().sniff(text[:8192], delimiters=",;\t|")
        except csv.Error:
            dialect = csv.excel
        rows = list(csv.reader(io.StringIO(text), dialect))
    kept = [r for r in rows if any((c or "").strip() for c in r)]
    EMPTY_ROWS[slot] = len(rows) - len(kept)
    rows = kept
    if len(rows) < 2:
        raise ValueError("The file needs a header row and at least one data row")
    taken = set()
    headers = [clean_name(h, taken) for h in rows[0]]
    n = len(headers)
    return headers, [(list(r) + [""] * n)[:n] for r in rows[1:]]


def load_pay(slot, headers, rows):
    """Payment files go straight into their own typed table; no bubbles, no calcs."""
    types = {h: detect_type([r[i] for r in rows]) for i, h in enumerate(headers)}
    DB.execute(f"DROP TABLE IF EXISTS {slot}")
    DB.execute(f"CREATE TABLE {slot} "
               f"({', '.join(f'{q(h)} {SQL_TYPE[types[h]]}' for h in headers)})")
    out = []
    for r in rows:
        vals = []
        for h, txt in zip(headers, r):
            t = types[h]
            v = txt if t == "string" else (PARSERS[t](txt) if (txt or "").strip() else None)
            vals.append(v.isoformat() if isinstance(v, date) else
                        (1 if v is True else 0 if v is False else v))
        out.append(vals)
    DB.executemany(f"INSERT INTO {slot} VALUES ({','.join('?' * len(headers))})", out)
    DB.commit()
    PAY[slot] = {"headers": headers, "types": types, "rows": len(out), "raw": rows}


@app.post("/upload")
def upload():
    global RAW, HEADERS, TYPES, CALCS
    f = request.files.get("file")
    slot = request.args.get("slot", "tape")
    if not f:
        return jsonify(error="No file"), 400
    if slot not in ("tape", "esperados", "recibidos"):
        return jsonify(error="Unknown slot"), 400
    try:
        headers, rows = read_table(f.filename or "", f.read(), slot)
    except ValueError as e:
        return jsonify(error=str(e)), 400
    except Exception as e:
        return jsonify(error=f"Could not read the file: {e}"), 400

    if slot != "tape":
        if DB is None:
            return jsonify(error="Upload the loan tape first"), 400
        load_pay(slot, headers, rows)
        return jsonify(slot=slot, rows=PAY[slot]["rows"], columns=headers)

    HEADERS, RAW, CALCS = headers, rows, []
    PAY.clear()
    TYPES = {h: detect_type([r[i] for r in RAW]) for i, h in enumerate(HEADERS)}
    rebuild()
    return jsonify(columns=bubbles(), rows=len(RAW))


@app.post("/retype")
def retype():
    col, new = request.json.get("column"), request.json.get("type")
    if col not in HEADERS or new not in PARSERS:
        return jsonify(error="Unknown column or type"), 400
    old = TYPES[col]
    TYPES[col] = new
    fails = rebuild()[col]
    samples = []
    if fails:
        i = HEADERS.index(col)
        samples = [r[i] for r in RAW
                   if (r[i] or "").strip() and PARSERS[new](r[i]) is None][:5]
    return jsonify(columns=bubbles(), failed=fails, total=len(RAW),
                   samples=samples, previous=old)


@app.post("/revert")
def revert():
    col, prev = request.json.get("column"), request.json.get("type")
    if col in HEADERS and prev in PARSERS:
        TYPES[col] = prev
        rebuild()
    return jsonify(columns=bubbles())


@app.post("/calc")
def calc():
    name = (request.json.get("name") or "").strip()
    expr = (request.json.get("expr") or "").strip()
    if not name or not expr:
        return jsonify(error="Name and expression are both required"), 400
    if ";" in expr or "--" in expr:
        return jsonify(error="One expression only: ; and -- are not allowed"), 400
    name = clean_name(name, set())
    if name in TYPES:
        return jsonify(error=f'"{name}" already exists'), 400
    try:
        DB.execute(f"SELECT ({expr}) FROM tape LIMIT 1").fetchone()
        apply_calc(name, expr)
    except sqlite3.Error as e:
        return jsonify(error=f"Invalid expression: {e}"), 400
    CALCS.append((name, expr))
    TYPES[name] = calc_type(name)
    return jsonify(columns=bubbles())


@app.post("/generate")
def generate():
    body = request.json
    saldo, dpd = body.get("saldo"), body.get("dpd")
    # min/max/average only mean something for numbers and dates
    summary = [c for c in (body.get("summary") or [])
               if TYPES.get(c) in ("integer", "float", "date")][:6]
    if saldo not in TYPES or dpd not in TYPES:
        return jsonify(error="Map both Saldo Insoluto and Días de Atraso"), 400
    if TYPES[saldo] not in ("float", "integer") or TYPES[dpd] not in ("float", "integer"):
        return jsonify(error="Saldo Insoluto and Días de Atraso must be numeric"), 400
    par = [c for c in (body.get("par") or []) if c in TYPES]
    hist = [c for c in (body.get("hist") or [])
            if TYPES.get(c) in ("integer", "float")]
    conc = [c for c in (body.get("conc") or []) if c in TYPES]
    coc = None
    if body.get("mode") == "coc":
        cid, orig = body.get("credit_id"), body.get("orig")
        odate = body.get("orig_date")
        missing = [n for n, v in (("Credit ID", cid), ("Originated", orig),
                                  ("Origination date", odate)) if v not in TYPES]
        if missing:
            return jsonify(error="Map " + ", ".join(missing)), 400
        if TYPES[odate] != "date":
            return jsonify(error=f"{odate} is {TYPES[odate]}, not a date"), 400
        if "recibidos" not in PAY or "esperados" not in PAY:
            return jsonify(error="COC needs both payment files uploaded"), 400
        coc = {"credit_id": cid, "orig": orig, "orig_date": odate,
               "dedupe": bool(body.get("dedupe", True)),
               "esp": pay_columns("esperados", cid),
               "rec": pay_columns("recibidos", cid)}
        for slot in ("esp", "rec"):
            m = coc[slot]
            if not m or not m["id"] or not m["amount"] or not m["date"]:
                return jsonify(error=f"Could not find the id, date and amount "
                                     f"columns in the {slot} file"), 400
    blanks = DB.execute(
        f"SELECT COUNT(*) FROM tape WHERE {q(dpd)} IS NULL").fetchone()[0]
    wide = [(c, distinct_values(c)[1]) for c in par]
    path, stats = build_workbook(saldo, dpd, summary, par, hist, coc, conc)
    fid = uuid.uuid4().hex
    FILES[fid] = path
    return jsonify(id=fid, blanks=blanks, total=len(RAW),
                   dropped=[c for c in (body.get("summary") or []) if c not in summary],
                   par=par, hist=hist, conc=conc,
                   wide=[c for c, n in wide if n > 300], **stats)


@app.get("/download/<fid>")
def download(fid):
    path = FILES.get(fid)
    if not path or not os.path.exists(path):
        return "Expired", 404
    return send_file(path, as_attachment=True,
                     download_name="Portfolio Analysis.xlsx")


# ---------- extra sheets ----------

PAR_DAYS = (7, 15, 30)
BAD_SHEET = re.compile(r"[\[\]:*?/\\]")
HIST_BINS = 12
OFF = 3            # three narrow spacer columns down the left of every sheet
MARGIN_W = 1.29
THEME = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Theme1.thmx")


def L(i):
    """Logical column number (1-based) to its letter on the sheet, past the margin."""
    return get_column_letter(i + OFF)


# brand palette, taken off the theme: dk2, accent4, accent5, accent2, accent3
INK, DEEP, MID, PALE, GREY, SAND = "1F1F1F", "0C473C", "4A655C", "9CC7B6", "6D6F6F", "D1C6BF"
# one colour scale and one data bar for the whole workbook. Capped at MID rather than
# DEEP so the cell text stays readable at the top of the ramp.
CF_LO, CF_MID, CF_HI = "FFE8F3EE", "FF" + PALE, "FF" + MID


def scale(ws, ref):
    """Pale-to-green ramp: darker means larger."""
    from openpyxl.formatting.rule import ColorScaleRule
    if ref:
        ws.conditional_formatting.add(ref, ColorScaleRule(
            start_type="min", start_color=CF_LO,
            mid_type="percentile", mid_value=50, mid_color=CF_MID,
            end_type="max", end_color=CF_HI))


def databar(ws, ref):
    from openpyxl.formatting.rule import DataBarRule
    if ref:
        ws.conditional_formatting.add(ref, DataBarRule(
            start_type="min", end_type="max", color="FF" + PALE))


def paint(ser, rgb, line=False, width=20000, dash=None):
    """Force one explicit colour on a series.

    Without this openpyxl leaves the series unpainted and Excel cycles the theme
    accents, which is why a single histogram came out black, grey, beige and two
    different greens."""
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    gp = GraphicalProperties()
    if line:
        gp.line = LineProperties(solidFill=rgb, w=width, prstDash=dash)
        gp.noFill = True
        ser.smooth = False
        from openpyxl.chart.marker import Marker
        ser.marker = Marker(symbol="none")
    else:
        gp.solidFill = rgb
        gp.line = LineProperties(noFill=True)
    ser.graphicalProperties = gp


def style_chart(ch, legend="b"):
    """Flat look: no chart border, hairline grey gridlines, legend out of the plot."""
    from openpyxl.chart.axis import ChartLines
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    ch.style = None
    ch.roundedCorners = False
    ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    hair = ChartLines(spPr=GraphicalProperties(
        ln=LineProperties(solidFill="EDEDED", w=6350)))
    axis_ln = GraphicalProperties(ln=LineProperties(solidFill="D8D8D8", w=6350))
    for ax in (ch.x_axis, ch.y_axis):
        ax.majorGridlines = None
        ax.spPr = axis_ln
        ax.majorTickMark = "none"
        ax.minorTickMark = "none"
    ch.y_axis.majorGridlines = hair
    if legend and ch.legend is not None:
        ch.legend.position = legend
        ch.legend.overlay = False
    ch.visible_cells_only = False


def margins(ws):
    for i in range(1, OFF + 1):
        ws.column_dimensions[get_column_letter(i)].width = MARGIN_W
    ws.sheet_view.showGridLines = False


def _lum(rgb):
    r, g, b = (int(rgb[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def theme_xml(fallback):
    """Theme1.thmx if it is sitting next to app.py, else the template's own theme.

    Either way Text 1 is forced dark. The supplied scheme sets dk1 to AFEAC1, a pale
    green, and Excel's Normal style is theme colour 1, so every default cell and every
    table style built on Text 1 renders near-invisible on the F6F6F6 background."""
    xml = fallback
    try:
        with zipfile.ZipFile(THEME) as z:
            name = next(n for n in z.namelist() if n.endswith("theme1.xml"))
            xml = z.read(name)
    except Exception:
        pass                      # ponytail: no theme file is not an error, use the template
    if not xml:
        return xml
    text = xml.decode("utf-8") if isinstance(xml, bytes) else xml
    m = re.search(r'(<a:dk1>\s*<a:srgbClr val=")([0-9A-Fa-f]{6})("\s*/>)', text)
    if m and _lum(m.group(2)) > 110:
        text = text[:m.start()] + m.group(1) + "1F1F1F" + m.group(3) + text[m.end():]
    return text.encode("utf-8")


def sheet_name(base, used):
    n = BAD_SHEET.sub("", base)[:31].strip() or "Sheet"
    base_n, i = n, 2
    while n.lower() in used:
        suffix = f" {i}"
        n, i = base_n[:31 - len(suffix)] + suffix, i + 1
    used.add(n.lower())
    return n


def header_row(ws, row, labels, widths=None):
    from openpyxl.styles import Alignment, Font, PatternFill
    for j, lab in enumerate(labels, start=1):
        c = ws.cell(row=row, column=j + OFF, value=lab)
        c.font = Font(name="Helvetica", size=8, color="FFF6F6F6", bold=True)
        c.fill = PatternFill("solid", fgColor="FF0C473C")
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    for j, w in enumerate(widths or [], start=1):
        ws.column_dimensions[L(j)].width = w


def title_cell(ws, text, sub=None):
    from openpyxl.styles import Font
    ws[f"{L(1)}1"] = text
    ws[f"{L(1)}1"].font = Font(name="Helvetica", size=10, bold=True)
    if sub:
        ws[f"{L(1)}2"] = sub
        ws[f"{L(1)}2"].font = Font(name="Helvetica", size=8, color="FF6D6F6F")


def distinct_values(col, limit=300):
    """Attribute values, most exposed first. Excel criteria treat * and ? as wildcards,
    so a value containing one would over-match; those are reported, not silently wrong."""
    rows = DB.execute(
        f"SELECT {q(col)}, COUNT(*) FROM tape WHERE {q(col)} IS NOT NULL "
        f"AND TRIM(CAST({q(col)} AS TEXT)) <> '' GROUP BY {q(col)} "
        f"ORDER BY COUNT(*) DESC").fetchall()
    return [r[0] for r in rows[:limit]], len(rows)


def block_title(ws, row, text, sub):
    """Title + subtitle for one stacked block. Left ungrouped so it stays visible
    when the block below it is collapsed."""
    from openpyxl.styles import Font
    ws.cell(row=row, column=(1) + OFF, value=text).font = Font(name="Helvetica", size=10, bold=True)
    ws.cell(row=row + 1, column=(1) + OFF, value=sub).font = Font(
        name="Helvetica", size=8, color="FF6D6F6F")


def collapse(ws, first, last):
    """Group first..last at outline level 1, left expanded. The +/- lands on the row
    after the group, which is why every block leaves a spacer row behind it."""
    for r in range(first, last + 1):
        ws.row_dimensions[r].outlineLevel = 1
        ws.row_dimensions[r].hidden = False     # grouped, but open on first open
    ws.sheet_properties.outlinePr.summaryBelow = True


def par_block(ws, top, attr, saldo, dpd):
    """One PAR block. Returns the row after it."""
    from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
    from openpyxl.styles import Font
    vals, total = distinct_values(attr)
    block_title(ws, top, f"Portfolio at Risk by {attr}",
                "PAR n = balance more than n days past due, over balance."
                + (f"  Showing {len(vals)} of {total} values." if total > len(vals) else ""))
    head = top + 3
    labels = [attr, "# Credits", "Outstanding"] \
        + [f"PAR {d} %" for d in PAR_DAYS] \
        + [f"PAR {d} Amount" for d in PAR_DAYS] \
        + [f"PAR {d} # " for d in PAR_DAYS]
    header_row(ws, head, labels, [26, 11, 16] + [10] * 3 + [16] * 3 + [10] * 3)

    first = head + 1
    for i, v in enumerate(vals):
        r = first + i
        ws.cell(row=r, column=(1) + OFF, value=v)
        a = f"${L(1)}{r}"
        ws.cell(row=r, column=(2) + OFF, value=f"=COUNTIFS(Data[{attr}],{a})")
        ws.cell(row=r, column=(3) + OFF, value=f"=SUMIFS(Data[{saldo}],Data[{attr}],{a})")
        for k, d in enumerate(PAR_DAYS):
            amt = L(7 + k)
            ws.cell(row=r, column=(7 + k) + OFF,
                    value=f'=SUMIFS(Data[{saldo}],Data[{attr}],{a},Data[{dpd}],">{d}")')
            ws.cell(row=r, column=(4 + k) + OFF,
                    value=f"=IFERROR({amt}{r}/${L(3)}{r},0)")
            ws.cell(row=r, column=(10 + k) + OFF,
                    value=f'=COUNTIFS(Data[{attr}],{a},Data[{dpd}],">{d}")')
    tot = first + len(vals)
    ws.cell(row=tot, column=(1) + OFF, value="TOTAL").font = Font(name="Helvetica", size=8, bold=True)
    ws.cell(row=tot, column=(2) + OFF, value=f"=SUM({L(2)}{first}:{L(2)}{tot - 1})")
    ws.cell(row=tot, column=(3) + OFF, value=f"=SUM({L(3)}{first}:{L(3)}{tot - 1})")
    for k in range(len(PAR_DAYS)):
        amt, cnt = L(7 + k), L(10 + k)
        ws.cell(row=tot, column=(7 + k) + OFF, value=f"=SUM({amt}{first}:{amt}{tot - 1})")
        ws.cell(row=tot, column=(10 + k) + OFF, value=f"=SUM({cnt}{first}:{cnt}{tot - 1})")
        ws.cell(row=tot, column=(4 + k) + OFF,
                value=f"=IFERROR({amt}{tot}/${L(3)}{tot},0)")

    for r in range(first, tot + 1):
        for col in (2,) + tuple(range(10, 13)):
            ws.cell(row=r, column=(col) + OFF).number_format = "#,##0"
        ws.cell(row=r, column=(3) + OFF).number_format = ACCT
        for k in range(len(PAR_DAYS)):
            ws.cell(row=r, column=(4 + k) + OFF).number_format = "0.00%"
            ws.cell(row=r, column=(7 + k) + OFF).number_format = ACCT
        for col in range(1, 13):
            f = ws.cell(row=r, column=(col) + OFF).font
            ws.cell(row=r, column=(col) + OFF).font = Font(name="Helvetica", size=8, bold=f.bold)

    # darker means more at risk; the data bar excludes TOTAL so one big row does
    # not flatten every other bar
    scale(ws, f"{L(4)}{first}:{L(6)}{tot}")
    databar(ws, f"{L(3)}{first}:{L(3)}{tot - 1}")
    collapse(ws, head, tot)
    return tot + 2


def add_par_sheets(wb, used, attrs, saldo, dpd):
    """Every breakout stacked on ONE sheet, each block collapsed to its title."""
    ws = wb.create_sheet(sheet_name("PAR", used))
    margins(ws)
    row = 1
    for attr in attrs:
        row = par_block(ws, row, attr, saldo, dpd)
    return ws


CONC_TOP = 20


def conc_block(ws, top, attr, saldo):
    """Top exposures by attribute, with cumulative share of the book."""
    from openpyxl.formatting.rule import DataBarRule
    from openpyxl.styles import Font
    rows = DB.execute(
        f"SELECT {q(attr)}, SUM({q(saldo)}) s FROM tape "
        f"WHERE {q(attr)} IS NOT NULL AND TRIM(CAST({q(attr)} AS TEXT)) <> '' "
        f"GROUP BY {q(attr)} ORDER BY s DESC").fetchall()
    vals = [r[0] for r in rows[:CONC_TOP]]
    rest = len(rows) - len(vals)
    block_title(ws, top, f"Concentration by {attr}",
                f"Largest {len(vals)} of {len(rows)} by outstanding balance, "
                f"cumulative share of the book.")
    head = top + 3
    header_row(ws, head, [attr, "# Credits", "Outstanding", "% of Book", "Cumulative %"],
               [30, 11, 18, 11, 13])
    first = head + 1
    for i, v in enumerate(vals):
        r = first + i
        ws.cell(row=r, column=(1) + OFF, value=v)
        ws.cell(row=r, column=(2) + OFF, value=f"=COUNTIFS(Data[{attr}],${L(1)}{r})")
        ws.cell(row=r, column=(3) + OFF,
                value=f"=SUMIFS(Data[{saldo}],Data[{attr}],${L(1)}{r})")
    last = first + len(vals) - 1
    other = last + 1
    tot = other + 1
    if rest:
        ws.cell(row=other, column=(1) + OFF, value=f"Other ({rest} values)")
        ws.cell(row=other, column=(2) + OFF,
                value=f"={L(2)}{tot}-SUM({L(2)}{first}:{L(2)}{last})")
        ws.cell(row=other, column=(3) + OFF,
                value=f"={L(3)}{tot}-SUM({L(3)}{first}:{L(3)}{last})")
    else:
        other = last
        tot = last + 1
    ws.cell(row=tot, column=(1) + OFF, value="TOTAL").font = Font(name="Helvetica", size=8, bold=True)
    ws.cell(row=tot, column=(2) + OFF, value=f"=COUNTA(Data[{attr}])")
    ws.cell(row=tot, column=(3) + OFF, value=f"=SUM(Data[{saldo}])")
    for r in range(first, tot + 1):
        ws.cell(row=r, column=(4) + OFF, value=f"=IFERROR({L(3)}{r}/${L(3)}${tot},0)")
        if r <= other:
            ws.cell(row=r, column=(5) + OFF,
                    value=f"=IFERROR(SUM(${L(3)}${first}:{L(3)}{r})/${L(3)}${tot},0)")
        ws.cell(row=r, column=(2) + OFF).number_format = "#,##0"
        ws.cell(row=r, column=(3) + OFF).number_format = ACCT
        ws.cell(row=r, column=(4) + OFF).number_format = "0.00%"
        ws.cell(row=r, column=(5) + OFF).number_format = "0.00%"
        for col in range(1, 6):
            f = ws.cell(row=r, column=(col) + OFF).font
            ws.cell(row=r, column=(col) + OFF).font = Font(name="Helvetica", size=8, bold=f.bold)
    databar(ws, f"{L(3)}{first}:{L(3)}{max(first, tot - 1)}")
    scale(ws, f"{L(4)}{first}:{L(5)}{max(first, tot - 1)}")
    collapse(ws, head, tot)
    return tot + 2


def add_concentration_sheet(wb, used, attrs, saldo):
    ws = wb.create_sheet(sheet_name("Concentration", used))
    margins(ws)
    row = 1
    for attr in attrs:
        row = conc_block(ws, row, attr, saldo)
    return ws


def add_histogram_sheet(wb, used, cols):
    """One binned frequency block + bar chart per selected numeric column.

    Bin edges are computed here from the data, but the counts are live COUNTIFS over
    the Data table, so editing the tape re-bins the heights (not the edges)."""
    from openpyxl.chart import BarChart, Reference
    from openpyxl.styles import Font
    ws = wb.create_sheet(sheet_name("Histograms", used))
    margins(ws)
    title_cell(ws, "Distributions",
               "One row per credit in the tape. Bin edges fixed at generation; counts are live.")
    for i, w in enumerate((22, 12, 12, 12), start=1):
        ws.column_dimensions[L(i)].width = w

    row = 4
    for c in cols:
        lo, hi = DB.execute(
            f"SELECT MIN({q(c)}), MAX({q(c)}) FROM tape WHERE {q(c)} IS NOT NULL").fetchone()
        if lo is None or hi is None or hi == lo:
            continue
        step = (hi - lo) / HIST_BINS
        ws.cell(row=row, column=(1) + OFF, value=c).font = Font(name="Helvetica", size=8, bold=True)
        header_row(ws, row + 1, ["Bin", "From", "To", "# Credits"])
        for b in range(HIST_BINS):
            a, z = lo + b * step, lo + (b + 1) * step
            r = row + 2 + b
            last = b == HIST_BINS - 1
            ws.cell(row=r, column=(1) + OFF,
                    value=f"{a:,.6g} to {z:,.6g}" + (" (incl.)" if last else ""))
            ws.cell(row=r, column=(2) + OFF, value=a)
            ws.cell(row=r, column=(3) + OFF, value=z)
            op = "<=" if last else "<"
            ws.cell(row=r, column=(4) + OFF,
                    value=f'=COUNTIFS(Data[{c}],">={a!r}",Data[{c}],"{op}{z!r}")')
            for col in (2, 3):
                ws.cell(row=r, column=(col) + OFF).number_format = "#,##0.00"
            ws.cell(row=r, column=(4) + OFF).number_format = "#,##0"
            for col in range(1, 5):
                ws.cell(row=r, column=(col) + OFF).font = Font(name="Helvetica", size=8)
        databar(ws, f"{L(4)}{row + 2}:{L(4)}{row + 1 + HIST_BINS}")
        ch = BarChart()
        ch.type, ch.title, ch.legend = "col", f"{c} distribution", None
        ch.height, ch.width = 6.2, 13
        ch.gapWidth = 40
        ch.add_data(Reference(ws, min_col=(4) + OFF, min_row=row + 1, max_row=row + 1 + HIST_BINS),
                    titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=(1) + OFF, min_row=row + 2, max_row=row + 1 + HIST_BINS))
        style_chart(ch, legend=None)
        paint(ch.series[0], DEEP)
        # the bin rows are collapsed below; without this the chart plots nothing,
        # because Excel omits hidden cells from a series by default. NOTE the openpyxl
        # attribute is visible_cells_only; setting plotVisOnly does nothing at all.
        ch.visible_cells_only = False
        ws.add_chart(ch, f"{L(6)}{row}")
        collapse(ws, row + 1, row + 1 + HIST_BINS)
        row += HIST_BINS + 4
    return ws


def pay_columns(slot, credit_id):
    """Work out which payment columns are the loan id, the date and the cash amount.

    The id is found by value overlap with the tape rather than by name, because the
    two files spell it differently ('CodPrestamo NB' vs 'Código de Préstamo')."""
    meta = PAY.get(slot)
    if not meta:
        return None
    heads, types = meta["headers"], meta["types"]
    ids = {str(r[0]) for r in DB.execute(
        f"SELECT DISTINCT {q(credit_id)} FROM tape WHERE {key_ok(credit_id)}")}
    best, best_hit = None, 0
    for h in heads:
        if types[h] not in ("integer", "float", "string"):
            continue
        vals = {str(r[0]) for r in DB.execute(
            f"SELECT DISTINCT {q(h)} FROM {slot} WHERE {q(h)} IS NOT NULL LIMIT 5000")}
        if not vals:
            continue
        hit = len(vals & ids) / len(vals)
        if hit > best_hit:
            best, best_hit = h, hit
    dates = [h for h in heads if types[h] == "date"]
    nums = [h for h in heads if types[h] in ("integer", "float") and h != best]
    amt = next((h for h in nums if re.search(r"total", h, re.I)), None)
    if amt is None and nums:
        sums = [(DB.execute(f"SELECT SUM({q(h)}) FROM {slot}").fetchone()[0] or 0, h)
                for h in nums]
        amt = max(sums)[1]
    return {"id": best, "id_hit": best_hit, "date": dates[0] if dates else None,
            "amount": amt, "headers": heads, "types": types}


def cohort_map(credit_id, orig_date):
    """credit id (as text) -> origination cohort as a YYYYMM integer."""
    out = {}
    for cid, first in DB.execute(
            f"SELECT CAST({q(credit_id)} AS TEXT), MIN({q(orig_date)}) FROM tape "
            f"WHERE {key_ok(credit_id)} GROUP BY 1"):
        c = (first or "")[:7].replace("-", "")
        if c.isdigit():
            out[cid] = int(c)
    return out


def add_pay_sheet(wb, used, slot, table, cols_meta, credit_id, dedupe, cmap=None):
    """Payment rows into their own sheet as an Excel Table, filtered to the tape.

    dedupe applies DISTINCT over every column. A schedule that was exported twice and
    stacked will otherwise double its expected cash. Never dedupe receipts: two identical
    receipt rows are usually two real payments of the same amount on the same day.

    With cmap, two key columns are appended: the credit's origination Cohort and MoB,
    the months from origination to that payment. The vintage triangle pivots on them.
    ponytail: written as values, not formulas, like every other key here. Editing a
    payment amount still updates the triangle; editing its date would not."""
    heads, types = cols_meta["headers"], cols_meta["types"]
    idc, datec = cols_meta["id"], cols_meta["date"]
    sel = ", ".join(q(h) for h in heads)
    sql = (f"SELECT {'DISTINCT ' if dedupe else ''}{sel} FROM {slot} "
           f"WHERE CAST({q(idc)} AS TEXT) IN {tape_ids(credit_id)}")
    ws = wb.create_sheet(sheet_name(table, used))
    margins(ws)
    if cmap is not None:
        heads = heads + ["Cohort", "MoB"]
        types = dict(types, Cohort="integer", MoB="integer")
    ws.append([None] * OFF + heads)
    fmt = {"date": "dd/mm/yyyy", "float": "#,##0.00", "integer": "#,##0"}
    ii, di = cols_meta["headers"].index(idc), cols_meta["headers"].index(datec)
    n = 0
    for r in DB.execute(sql):
        vals = [date.fromisoformat(v) if (types[h] == "date" and v) else v
                for h, v in zip(cols_meta["headers"], r)]
        if cmap is not None:
            coh = cmap.get(str(r[ii]))
            d = vals[di]
            vals += [coh, (d.year * 12 + d.month) - (coh // 100 * 12 + coh % 100)
                     if coh and isinstance(d, date) else None]
        ws.append([None] * OFF + vals)
        n += 1
    if n == 0:
        return ws, 0
    for i, h in enumerate(heads, start=1):
        nf = fmt.get(types[h])
        if nf:
            for c in ws.iter_rows(min_row=2, min_col=(i) + OFF, max_col=(i) + OFF):
                c[0].number_format = nf
        ws.column_dimensions[L(i)].width = max(10, min(26, len(h) + 4))
    t = Table(displayName=table, ref=f"{L(1)}1:{L(len(heads))}{n + 1}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(t)
    ws.freeze_panes = f"{L(1)}2"
    return ws, n


def add_credits_sheet(wb, used, credit_id, saldo, orig, orig_date, rec_meta):
    """One row per credit id. Every metric is a live SUMIFS; only the keys are values.

    Payments are keyed to the credit id, so the tape is rolled up to that grain before
    joining. Where a tape holds one row per credit the roll-up is 1:1 and changes
    nothing; where it holds several, it is what stops one payment being counted once
    per row. No assumption about why a credit spans rows."""
    from openpyxl.styles import Font
    rows = DB.execute(
        f"SELECT CAST({q(credit_id)} AS TEXT), MIN({q(orig_date)}) "
        f"FROM tape WHERE {key_ok(credit_id)} "
        f"GROUP BY CAST({q(credit_id)} AS TEXT) ORDER BY 2, 1").fetchall()
    ws = wb.create_sheet(sheet_name("Credits", used))
    margins(ws)
    heads = ["Credit ID", "Cohort", "Rows", "Originated", "Received",
             "Outstanding", "MOIC", "Closed"]
    ws.append([None] * OFF + heads)
    ridc, ramt = rec_meta["id"], rec_meta["amount"]
    for gid, first in rows:
        # numeric YYYYMM: the MOIC sheet matches on it with SUMIFS, and a text
        # "202304" will not reliably match a numeric 202304 criterion
        c = (first or "")[:7].replace("-", "")
        ws.append([None] * OFF + [gid, int(c) if c.isdigit() else None,
                                  None, None, None, None, None, None])
    n = len(rows)
    for i in range(2, n + 2):
        a = f"${L(1)}{i}"
        ws.cell(row=i, column=(3) + OFF, value=f"=COUNTIFS(Data[{credit_id}],{a})")
        ws.cell(row=i, column=(4) + OFF, value=f"=SUMIFS(Data[{orig}],Data[{credit_id}],{a})")
        ws.cell(row=i, column=(5) + OFF,
                value=f"=SUMIFS(Recibidos[{ramt}],Recibidos[{ridc}],{a})")
        ws.cell(row=i, column=(6) + OFF, value=f"=SUMIFS(Data[{saldo}],Data[{credit_id}],{a})")
        ws.cell(row=i, column=(7) + OFF, value=f"=IFERROR({L(5)}{i}/{L(4)}{i},0)")
        ws.cell(row=i, column=(8) + OFF, value=f"=IF({L(6)}{i}<=0,1,0)")
        for col, nf in ((3, "#,##0"), (4, ACCT), (5, ACCT), (6, ACCT),
                        (7, "0.00x"), (8, "0")):
            ws.cell(row=i, column=(col) + OFF).number_format = nf
    header_row(ws, 1, heads, [14, 10, 10, 16, 16, 16, 10, 8])
    ws.freeze_panes = f"{L(1)}2"
    if n:
        t = Table(displayName="Credits", ref=f"{L(1)}1:{L(8)}{n + 1}")
        t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
        ws.add_table(t)
    return ws, n


MOIC_BINS = 20


def add_moic_sheet(wb, used, orig_date):
    """Cohort MOIC progression plus the 0.0-2.0x distribution, both off Credits."""
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.styles import Font
    ws = wb.create_sheet(sheet_name("MOIC", used))
    margins(ws)
    title_cell(ws, "MOIC",
               "Cash received over amount originated, per credit. Cohort MOIC is "
               "summed-received over summed-originated, so it is value weighted.")
    cohorts = [r[0] for r in DB.execute(
        f"SELECT DISTINCT SUBSTR({q(orig_date)},1,7) c FROM tape "
        f"WHERE {q(orig_date)} IS NOT NULL ORDER BY c")]
    cohorts = [c.replace("-", "") for c in cohorts]

    header_row(ws, 4, ["Cohort", "Originated", "Received", "Cohort MOIC", "1.0x"],
               [12, 18, 18, 13, 8])
    for i, c in enumerate(cohorts):
        r = 5 + i
        ws.cell(row=r, column=(1) + OFF, value=int(c))
        ws.cell(row=r, column=(2) + OFF,
                value=f"=SUMIFS(Credits[Originated],Credits[Cohort],${L(1)}{r})")
        ws.cell(row=r, column=(3) + OFF,
                value=f"=SUMIFS(Credits[Received],Credits[Cohort],${L(1)}{r})")
        ws.cell(row=r, column=(4) + OFF, value=f"=IFERROR({L(3)}{r}/{L(2)}{r},0)")
        ws.cell(row=r, column=(5) + OFF, value=1)
        ws.cell(row=r, column=(2) + OFF).number_format = ACCT
        ws.cell(row=r, column=(3) + OFF).number_format = ACCT
        ws.cell(row=r, column=(4) + OFF).number_format = "0.00x"
        ws.cell(row=r, column=(5) + OFF).number_format = "0.0x"
    last = 4 + len(cohorts)
    if cohorts:
        databar(ws, f"{L(2)}5:{L(3)}{last}")
        scale(ws, f"{L(4)}5:{L(4)}{last}")

    if cohorts:
        bar = BarChart()
        bar.type, bar.title = "col", "Rentability | Progression"
        bar.height, bar.width = 8, 22
        bar.gapWidth = 60
        bar.add_data(Reference(ws, min_col=(2) + OFF, max_col=(3) + OFF, min_row=4, max_row=last),
                     titles_from_data=True)
        bar.set_categories(Reference(ws, min_col=(1) + OFF, min_row=5, max_row=last))
        ln = LineChart()
        ln.add_data(Reference(ws, min_col=(4) + OFF, max_col=(5) + OFF, min_row=4, max_row=last),
                    titles_from_data=True)
        ln.y_axis.axId = 200
        ln.y_axis.title = "MOIC"
        bar.y_axis.crosses = "autoZero"
        style_chart(bar)
        paint(bar.series[0], DEEP)          # Originated
        paint(bar.series[1], PALE)          # Received
        # "+=" keeps the line series on ln, it does not fold them into bar.series
        paint(ln.series[0], INK, line=True, width=22000)                  # Cohort MOIC
        paint(ln.series[1], SAND, line=True, width=12700, dash="dash")    # 1.0x
        bar += ln
        ws.add_chart(bar, f"{L(7)}4")

    # distribution: differenced CDF, same construction as the reference model
    top = last + 3
    ws.cell(row=top, column=(1) + OFF, value="Distribution").font = Font(
        name="Helvetica", size=8, bold=True)
    header_row(ws, top + 1, ["From", "To", "All", "Closed", "Active"])
    for b in range(MOIC_BINS):
        r, lo, hi = top + 2 + b, b * 0.1, (b + 1) * 0.1
        ws.cell(row=r, column=(1) + OFF, value=round(lo, 2))
        ws.cell(row=r, column=(2) + OFF, value=round(hi, 2))
        prev = lambda c: "" if b == 0 else f"-SUM({c}{top + 2}:{c}{r - 1})"
        ws.cell(row=r, column=(3) + OFF,
                value=f'=COUNTIFS(Credits[MOIC],"<="&${L(2)}{r}){prev(L(3))}')
        ws.cell(row=r, column=(4) + OFF,
                value=f'=COUNTIFS(Credits[Closed],1,Credits[MOIC],"<="&${L(2)}{r}){prev(L(4))}')
        ws.cell(row=r, column=(5) + OFF,
                value=f'=COUNTIFS(Credits[Closed],0,Credits[MOIC],"<="&${L(2)}{r}){prev(L(5))}')
        for col in (1, 2):
            ws.cell(row=r, column=(col) + OFF).number_format = "0.0x"
        for col in (3, 4, 5):
            ws.cell(row=r, column=(col) + OFF).number_format = "#,##0"
    tail = top + 1 + MOIC_BINS
    over = tail + 1
    ws.cell(row=over, column=(1) + OFF, value="above 2.0x")
    ws.cell(row=over, column=(3) + OFF, value=f'=COUNTIFS(Credits[MOIC],">2")')
    ws.cell(row=over, column=(4) + OFF, value=f'=COUNTIFS(Credits[Closed],1,Credits[MOIC],">2")')
    ws.cell(row=over, column=(5) + OFF, value=f'=COUNTIFS(Credits[Closed],0,Credits[MOIC],">2")')

    scale(ws, f"{L(3)}{top + 2}:{L(5)}{tail}")

    dist = LineChart()
    dist.title = "Rentability | Distribution"
    dist.height, dist.width = 8, 22
    dist.add_data(Reference(ws, min_col=(3) + OFF, max_col=(5) + OFF, min_row=top + 1, max_row=tail),
                  titles_from_data=True)
    dist.set_categories(Reference(ws, min_col=(2) + OFF, min_row=top + 2, max_row=tail))
    style_chart(dist)
    for ser, rgb in zip(dist.series, (INK, DEEP, PALE)):    # All / Closed / Active
        paint(ser, rgb, line=True, width=20000)
    ws.add_chart(dist, f"{L(7)}{top}")
    return ws


def coc_triangle(ws, top, rec_meta, orig_date):
    """Cumulative MOIC by origination cohort (down) and months on book (across).

    Cell (cohort, n) = cash received from that cohort through month n on book, over
    what the cohort originated. Each row stops at the months actually elapsed, which
    is what gives the triangle its shape rather than a wall of zeros."""
    from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
    from openpyxl.styles import Alignment, Font
    ra = rec_meta["amount"]
    cohorts = [r[0] for r in DB.execute(
        f"SELECT DISTINCT SUBSTR({q(orig_date)},1,7) c FROM tape "
        f"WHERE {q(orig_date)} IS NOT NULL ORDER BY c")]
    cohorts = [int(c.replace("-", "")) for c in cohorts if c.replace("-", "").isdigit()]
    asof = DB.execute(
        f"SELECT MAX({q(rec_meta['date'])}) FROM recibidos").fetchone()[0] or ""
    if not cohorts or len(asof) < 7:
        return top
    a = int(asof[:4]) * 12 + int(asof[5:7])
    span = {c: a - (c // 100 * 12 + c % 100) for c in cohorts}
    n_max = max(max(span.values()), 0)

    block_title(ws, top, "Evolución de MOICs",
                "Cumulative cash received over amount originated, by cohort and "
                "months on book.")
    head = top + 3
    ws.cell(row=head - 1, column=(4 + n_max // 2) + OFF, value="MoB").font = Font(
        name="Helvetica", size=8, bold=True)
    header_row(ws, head, ["Month of Cohort", "Disbursed", "# Credits"]
               + list(range(n_max + 1)), [17, 16, 10] + [7] * (n_max + 1))
    first = head + 1
    for i, c in enumerate(cohorts):
        r = first + i
        lab = ws.cell(row=r, column=(1) + OFF, value=date(c // 100, c % 100, 1))
        lab.number_format = "mmmm yyyy"
        lab.font = Font(name="Helvetica", size=8, bold=True)
        # the label is a real date so Excel can render "January 2025", but the key
        # stored on Recibidos and Credits is the YYYYMM integer, so rebuild it here
        key = f"(YEAR(${L(1)}{r})*100+MONTH(${L(1)}{r}))"
        size = ws.cell(row=r, column=(2) + OFF,
                       value=f"=SUMIFS(Credits[Originated],Credits[Cohort],{key})")
        size.number_format = ACCT
        cnt = ws.cell(row=r, column=(3) + OFF,
                      value=f"=COUNTIFS(Credits[Cohort],{key})")
        cnt.number_format = "#,##0"
        for cell in (size, cnt):
            cell.font = Font(name="Helvetica", size=8)
        for n in range(min(span[c], n_max) + 1):
            # denominator is the Disbursed cell on this row, not a second SUMIFS
            cell = ws.cell(row=r, column=(4 + n) + OFF, value=(
                f'=IFERROR(SUMIFS(Recibidos[{ra}],Recibidos[Cohort],{key},'
                f'Recibidos[MoB],"<={n}")/${L(2)}{r},"")'))
            cell.number_format = '0.00"x"'
            cell.font = Font(name="Helvetica", size=8)
            cell.alignment = Alignment(horizontal="center")
    last = first + len(cohorts) - 1
    scale(ws, f"{L(4)}{first}:{L(4 + n_max)}{last}")
    databar(ws, f"{L(2)}{first}:{L(2)}{last}")
    databar(ws, f"{L(3)}{first}:{L(3)}{last}")
    ws.freeze_panes = f"{L(4)}{first}"
    return last + 2


def add_coc_sheet(wb, used, esp_meta, rec_meta, esp_n, rec_n, orig_date):
    """Vintage MOIC triangle, with the calendar-month collection table below it."""
    from openpyxl.chart import BarChart, LineChart, Reference
    months = sorted({r[0] for r in DB.execute(
        f"SELECT DISTINCT SUBSTR({q(esp_meta['date'])},1,7) FROM esperados "
        f"WHERE {q(esp_meta['date'])} IS NOT NULL")}
        | {r[0] for r in DB.execute(
            f"SELECT DISTINCT SUBSTR({q(rec_meta['date'])},1,7) FROM recibidos "
            f"WHERE {q(rec_meta['date'])} IS NOT NULL")})
    ws = wb.create_sheet(sheet_name("COC", used))
    margins(ws)
    tbl = coc_triangle(ws, 1, rec_meta, orig_date)

    block_title(ws, tbl, "Collections by month",
                "Scheduled against actual, by calendar month.")
    hrow = tbl + 3
    header_row(ws, hrow, ["Month", "Expected", "Received", "Collection %",
                          "Cum. Expected", "Cum. Received", "Cum. %"])
    ed, ea = esp_meta["date"], esp_meta["amount"]
    rd, ra = rec_meta["date"], rec_meta["amount"]
    base = hrow + 1
    for i, m in enumerate(months):
        r = base + i
        y, mo = int(m[:4]), int(m[5:7])
        ws.cell(row=r, column=(1) + OFF, value=date(y, mo, 1)).number_format = "mmm-yy"
        ws.cell(row=r, column=(2) + OFF, value=(
            f'=SUMIFS(Esperados[{ea}],Esperados[{ed}],">="&${L(1)}{r},'
            f'Esperados[{ed}],"<="&EOMONTH(${L(1)}{r},0))'))
        ws.cell(row=r, column=(3) + OFF, value=(
            f'=SUMIFS(Recibidos[{ra}],Recibidos[{rd}],">="&${L(1)}{r},'
            f'Recibidos[{rd}],"<="&EOMONTH(${L(1)}{r},0))'))
        ws.cell(row=r, column=(4) + OFF, value=f"=IFERROR({L(3)}{r}/{L(2)}{r},0)")
        ws.cell(row=r, column=(5) + OFF, value=f"=SUM(${L(2)}${base}:{L(2)}{r})")
        ws.cell(row=r, column=(6) + OFF, value=f"=SUM(${L(3)}${base}:{L(3)}{r})")
        ws.cell(row=r, column=(7) + OFF, value=f"=IFERROR({L(6)}{r}/{L(5)}{r},0)")
        for col in (2, 3, 5, 6):
            ws.cell(row=r, column=(col) + OFF).number_format = ACCT
        for col in (4, 7):
            ws.cell(row=r, column=(col) + OFF).number_format = "0.00%"
    last = hrow + len(months)
    if months:
        databar(ws, f"{L(2)}{base}:{L(3)}{last}")
        scale(ws, f"{L(4)}{base}:{L(4)}{last}")
        scale(ws, f"{L(7)}{base}:{L(7)}{last}")
    if months:
        bar = BarChart()
        bar.type, bar.title = "col", "Collections | Expected vs Received"
        bar.height, bar.width = 8, 22
        bar.gapWidth = 60
        bar.add_data(Reference(ws, min_col=(2) + OFF, max_col=(3) + OFF,
                               min_row=hrow, max_row=last), titles_from_data=True)
        bar.set_categories(Reference(ws, min_col=(1) + OFF, min_row=base, max_row=last))
        ln = LineChart()
        ln.add_data(Reference(ws, min_col=(7) + OFF, min_row=hrow, max_row=last),
                    titles_from_data=True)
        ln.y_axis.axId = 200
        ln.y_axis.title = "Cumulative %"
        style_chart(bar)
        paint(bar.series[0], PALE)          # Expected
        paint(bar.series[1], DEEP)          # Received
        paint(ln.series[0], INK, line=True, width=22000)     # Cumulative %
        bar += ln
        ws.add_chart(bar, f"{L(9)}{hrow}")
    collapse(ws, hrow, last)
    return ws


# ---------- data audit ----------

TYPE_ES = {"string": "texto", "integer": "entero", "float": "decimal",
           "date": "fecha", "boolean": "booleano"}
A1, A2, A3 = "1. Vacíos y nulos", "2. Llenado y capacidad", "3. Integridad y coherencia"
# assumptions the audit tests against; written as blue inputs so they can be tuned
SUPUESTOS = [
    ("rows", "Límite de filas por hoja de Excel", 1048576, "#,##0"),
    ("prec", "Precisión numérica de Excel (dígitos)", 15, "0"),
    ("alert", "Alerta de capacidad", 0.80, "0.00%"),
    ("part", "Umbral para particionar tablas de pagos", 0.25, "0.00%"),
    ("card", "Cardinalidad máxima para catálogo", 0.01, "0.00%"),
    ("idcard", "Cardinalidad mínima de un identificador", 0.99, "0.00%"),
    ("lowuse", "Uso bajo de VARCHAR(255)", 0.10, "0.00%"),
    ("varchar", "Límite VARCHAR por defecto (caracteres)", 255, "#,##0"),
    ("decimal", "Dígitos enteros de DECIMAL(18,2)", 16, "0"),
    ("smallint", "Máximo de SMALLINT", 32767, "#,##0"),
    ("int", "Máximo de INT", 2147483647, "#,##0"),
]
# smallest length that moves a text column up to the next declared size
LADDER = [(0, "CHAR(1)"), (2, "VARCHAR(16)"), (17, "VARCHAR(32)"), (33, "VARCHAR(64)"),
          (65, "VARCHAR(128)"), (129, "VARCHAR(255)"), (256, "TEXT")]


def case_variants(col):
    """Text values that differ only in capitalisation, e.g. 'TOYOTA' and 'Toyota'.

    Counted here rather than in the workbook because Excel's UNIQUE ignores case, so
    the live Distintos column merges exactly the variants this is meant to catch."""
    vals = {str(r[0]).strip() for r in DB.execute(
        f"SELECT DISTINCT {q(col)} FROM tape WHERE {q(col)} IS NOT NULL")}
    vals.discard("")
    return len(vals) - len({v.casefold() for v in vals})


def add_audit_sheet(wb, used, saldo, dpd, coc, pst):
    """Formula-driven data-quality audit: nulls, fill and capacity, integrity.

    A CSV or xlsx tape declares no NOT NULL or VARCHAR limits, so the schema is
    inferred from the data: usage is measured against what a naive import would
    declare and against Excel's own ceilings, and each column gets the constraint and
    type it should carry. Every metric is live over the workbook's tables; the few
    counts that only exist at load time (rows dropped before they reached the
    workbook, case variants Excel cannot see) are written as blue values."""
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Border, Font, PatternFill, Side
    from openpyxl.worksheet.formula import ArrayFormula

    ws = wb.create_sheet(sheet_name("Auditoría de Datos", used))
    margins(ws)
    title_cell(ws, "Auditoría de Datos",
               "Perfil de la cinta y de los archivos de pagos. Todo el cálculo es en vivo "
               "sobre las tablas del libro; los conteos de la carga van en azul.")
    for i, w in enumerate((22, 44, 12, 12, 22, 12, 12, 12, 10, 11, 11, 10, 10, 16,
                           14, 12, 18, 42), start=1):
        ws.column_dimensions[L(i)].width = w

    def put(r, c, v, fmt=None, bold=False, blue=False):
        if isinstance(v, tuple) and v[0] == "arr":      # single-cell array formula
            cell = ws.cell(row=r, column=c + OFF)
            cell.value = ArrayFormula(f"{L(c)}{r}", v[1])
        else:
            cell = ws.cell(row=r, column=c + OFF, value=v)
        cell.font = Font(name="Helvetica", size=8, bold=bold,
                         color="FF0432FF" if blue else None)
        if blue:                                        # house style for hardcodes
            cell.fill = PatternFill("solid", fgColor="FFFCFFBD")
            thin = Side(style="thin", color="FFBFBFBF")
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        if fmt:
            cell.number_format = fmt
        return cell

    # ---- what gets profiled ----
    prof = [("Data", c, TYPES[c]) for c in all_columns()]
    crit = {("Data", saldo), ("Data", dpd)}
    has = {"rec": False, "esp": False, "cred": False}
    if coc:
        crit |= {("Data", coc["credit_id"]), ("Data", coc["orig"]), ("Data", coc["orig_date"])}
        has = {"rec": pst.get("rec_n", 0) > 0, "esp": pst.get("esp_n", 0) > 0,
               "cred": pst.get("credits", 0) > 0}
        for key, tbl, extra in (("esp", "Esperados", []),
                                ("rec", "Recibidos", [("Cohort", "integer"), ("MoB", "integer")])):
            if not has[key]:                # no table in the workbook, nothing to reference
                continue
            m = coc[key]
            prof += [(tbl, h, m["types"][h]) for h in m["headers"]]
            prof += [(tbl, h, t) for h, t in extra]
            crit |= {(tbl, m["id"]), (tbl, m["date"]), (tbl, m["amount"])}
    first_col = all_columns()[0]
    variants = sum(case_variants(c) for c in all_columns() if TYPES[c] == "string")
    empty_t = EMPTY_ROWS.get("tape", 0)

    # ---- controls: (area, label, result, fmt, threshold key or None, severity, detail) ----
    ctrls = []
    add = ctrls.append
    if coc:
        empty_e, empty_r = EMPTY_ROWS.get("esperados", 0), EMPTY_ROWS.get("recibidos", 0)
        add((A1, "Registros completamente vacíos descartados en la carga",
             ("val", empty_t + empty_e + empty_r), "#,##0", None, "REVISAR",
             f"Cinta: {empty_t:,}. Esperados: {empty_e:,}. Recibidos: {empty_r:,}. "
             f"Se eliminan antes de crear las hojas."))
    else:
        add((A1, "Registros completamente vacíos descartados en la carga",
             ("val", empty_t), "#,##0", None, "REVISAR",
             "La carga elimina las filas sin ningún valor antes de crear la hoja Data."))
    add((A1, f"Registros sin {saldo} ni {dpd}",
         f'=COUNTIFS(Data[{saldo}],"",Data[{dpd}],"")', "#,##0", None, "RIESGO",
         "Filas sin los dos campos que definen el saldo en riesgo."))
    add((A1, f"Nulos en {saldo} (Saldo Insoluto)", ("nul", "Data", saldo), "#,##0", None,
         "RIESGO", "Campo crítico: sin saldo el crédito queda fuera de PAR y de los buckets."))
    add((A1, f"Nulos en {dpd} (Días de Atraso)", ("nul", "Data", dpd), "#,##0", None, "REVISAR",
         "Se tratan como al corriente; confirmar que así los reporta el acreditado."))
    add((A1, "Columnas completamente vacías", ("empty",), "#,##0", None, "REVISAR",
         "Columnas sin ningún valor: ocupan espacio sin aportar información."))
    add((A2, "Filas de Data frente al límite de Excel",
         ("ratio", f"ROWS(Data[{first_col}])"), "0.00%", "alert", "RIESGO",
         "Una hoja de Excel admite un número fijo de filas; ver Supuestos."))
    add((A2, "Columnas numéricas con más dígitos que la precisión de Excel", ("prec",),
         "#,##0", None, "RIESGO",
         "Excel conserva 15 dígitos significativos y redondea el resto sin avisar."))
    add((A2, "Columnas enteras que exceden el rango de INT", ("intover",), "#,##0", None, "RIESGO",
         "Requieren BIGINT en la base destino."))
    add((A2, "Columnas de texto que exceden VARCHAR(255)", ("textover",), "#,##0", None, "RIESGO",
         "Se truncarían al cargarlas con el tipo por defecto; requieren TEXT."))
    add((A2, "Columnas cerca de su límite de tipo", ("near",), "#,##0", None, "REVISAR",
         "Usan al menos la alerta de capacidad de su límite por defecto sin excederlo."))
    add((A3, "Filas duplicadas en la cinta",
         ("arr", f"=ROWS(Data[{first_col}])-ROWS(_xlfn.UNIQUE(Data[#Data]))"), "#,##0", None, "RIESGO",
         "Filas idénticas en todas las columnas: su saldo se cuenta dos veces."))
    add((A3, "Identificadores enteros con valores repetidos", ("iddup",), "#,##0", None, "REVISAR",
         "Columnas casi únicas por fila (probables identificadores) con algún valor repetido."))
    add((A3, "Valores de texto que solo difieren en mayúsculas", ("val", variants), "#,##0",
         None, "REVISAR",
         "Errores de captura: 'TOYOTA' y 'Toyota' cuentan como uno en Distintos."))
    add((A3, f"Saldos negativos en {saldo}", f'=COUNTIFS(Data[{saldo}],"<0")', "#,##0",
         None, "RIESGO", "Un saldo insoluto no puede ser negativo."))
    add((A3, f"Días de atraso negativos en {dpd}", f'=COUNTIFS(Data[{dpd}],"<0")', "#,##0",
         None, "RIESGO", "Un atraso negativo indica un error de fecha de corte."))
    add((A3, "Créditos liquidados con atraso", f'=COUNTIFS(Data[{saldo}],0,Data[{dpd}],">0")',
         "#,##0", None, "REVISAR",
         "Saldo en cero con días de atraso: estado incoherente del crédito."))
    if coc:
        cid, rec, esp = coc["credit_id"], coc["rec"], coc["esp"]
        add((A1, f"Nulos en {cid} (ID de crédito)", ("nul", "Data", cid), "#,##0", None, "RIESGO",
             "Sin ID el crédito no puede unirse a sus pagos."))
        add((A1, f"Nulos en {coc['orig_date']} (fecha de originación)",
             ("nul", "Data", coc["orig_date"]), "#,##0", None, "RIESGO",
             "Sin fecha el crédito queda fuera de la cosecha."))
        add((A1, f"Nulos en {coc['orig']} (monto originado)", ("nul", "Data", coc["orig"]),
             "#,##0", None, "RIESGO", "Es el denominador del MOIC."))
        add((A1, "Pagos sin ID de crédito",
             ("val", pst.get("esp_null_ids", 0) + pst.get("rec_null_ids", 0)), "#,##0", None,
             "RIESGO", f"Esperados: {pst.get('esp_null_ids', 0):,}. Recibidos: "
             f"{pst.get('rec_null_ids', 0):,}. No pueden asignarse a ningún crédito."))
        add((A3, "Filas que repiten un ID de crédito (exceso sobre IDs únicos)",
             ("dupid", cid), "#,##0", None, "REVISAR",
             "Un mismo ID aparece en varias filas; COC y MOIC agregan a nivel crédito."))
        if has["esp"] or has["rec"]:
            add((A3, "Llave de crédito con tipos distintos entre tablas", ("keytype",), "0",
                 None, "RIESGO",
                 "1 si el ID no tiene el mismo tipo en la cinta y en los pagos: la llave "
                 "foránea no puede declararse sin conversión."))
        if has["rec"]:
            add((A1, "Pagos recibidos sin fecha o sin monto",
                 ("nul2", "Recibidos", rec["date"], rec["amount"]), "#,##0", None, "RIESGO",
                 "Un pago sin fecha o sin monto no entra en la cosecha."))
            add((A2, "Filas de Recibidos frente al límite de Excel",
                 ("ratio", f"ROWS(Recibidos[{rec['id']}])"), "0.00%", "alert", "RIESGO",
                 "El archivo de pagos crece cada mes."))
            add((A3, "Pagos recibidos anteriores al mes de originación",
                 '=COUNTIFS(Recibidos[MoB],"<0")', "#,##0", None, "RIESGO",
                 "Pago fechado en un mes previo al desembolso: error de fecha o de ID."))
            add((A3, "Pagos recibidos con monto negativo",
                 f'=COUNTIFS(Recibidos[{rec["amount"]}],"<0")', "#,##0", None, "REVISAR",
                 "Reversos: se netean, no se eliminan."))
            add((A3, "Pagos recibidos huérfanos (ID fuera de la cinta)",
                 ("val", pst.get("rec_orphans", 0)), "#,##0", None, "RIESGO",
                 "Filas con ID que no existe en la cinta; se excluyeron del libro."))
        if has["esp"]:
            add((A2, "Filas de Esperados frente al límite de Excel",
                 ("ratio", f"ROWS(Esperados[{esp['id']}])"), "0.00%", "alert", "RIESGO",
                 "El calendario de pagos también crece."))
            add((A3, "Pagos esperados con ID fuera de la cinta",
                 ("val", pst.get("esp_orphans", 0)), "#,##0", None, "REVISAR",
                 "El calendario cubre más créditos que la cinta; se excluyeron del libro."))
            if coc["dedupe"]:
                add((A3, "Duplicados eliminados de pagos esperados en la carga",
                     ("val", pst.get("esp_dups", 0)), "#,##0", None, "REVISAR",
                     "Copias exactas de créditos de la cinta; ya no están en la hoja Esperados."))
            else:
                add((A3, "Filas duplicadas en pagos esperados",
                     ("val", pst.get("esp_dups", 0)), "#,##0", None, "RIESGO",
                     "Copias exactas que siguen en la hoja Esperados e inflan el flujo esperado."))
        if has["cred"]:
            add((A3, "Créditos sin ningún pago recibido", '=COUNTIFS(Credits[Received],0)',
                 "#,##0", None, "REVISAR",
                 "Huérfanos del lado de pagos; normal en originaciones recientes."))
            add((A3, "Créditos con saldo mayor al monto originado",
                 "=SUMPRODUCT(--(Credits[Outstanding]>Credits[Originated]))", "#,##0",
                 None, "RIESGO", "El saldo no puede superar lo desembolsado."))
            if has["esp"]:
                add((A3, "Créditos sin calendario de pagos esperados",
                     f"=SUMPRODUCT(--(COUNTIFS(Esperados[{esp['id']}],Credits[Credit ID])=0))",
                     "#,##0", None, "REVISAR",
                     "Su flujo esperado cuenta como cero en COC."))

    # ---- optimisations: (area, label, metric, fmt, applies-when, evidence template) ----
    # templates use {m} for this row's metric cell; never a literal column letter,
    # because the OFF margin moves every column
    opts = [
        (A2, "Declarar tipos ajustados en lugar de VARCHAR(255) genérico", ("lowuse",), "#,##0",
         ">0", '={m}&" de "&{ntext}&" columnas de texto usan menos de "&TEXT({s_lowuse},"0%")'
               '&" de VARCHAR(255). El tipo recomendado está en el perfil."'),
        (A3, "Normalizar columnas de baja cardinalidad en catálogos", ("catalog",), "#,##0",
         ">0", '={m}&" columnas de texto repiten pocos valores; guardarlas como llave a un '
               'catálogo reduce espacio y errores de captura."'),
        (A2, "Guardar los identificadores como texto", ("idlike",), "#,##0", ">0",
         '={m}&" columnas enteras son casi únicas por fila (probables identificadores); como '
         'número pierden ceros a la izquierda y quedan limitadas a la precisión de Excel."'),
        (A1, "Eliminar de la carga las columnas vacías", ("empty",), "#,##0", ">0",
         '={m}&" columnas no tienen ningún valor."'),
        (A1, "Declarar NOT NULL en los campos críticos", ("critok",), "#,##0", ">0",
         '={m}&" de "&{ncrit}&" campos críticos no tienen nulos y pueden declararse NOT NULL."'),
    ]
    if coc:
        opts.append((A3, "Separar la tabla de créditos de la tabla de filas",
                     ("dupid", coc["credit_id"]), "#,##0", ">0",
                     '={m}&" filas repiten un ID de crédito; una tabla por crédito y otra por '
                     'fila evita repetir los datos del crédito."'))
        if has["rec"]:
            opts.append((A2, "Indexar los pagos por ID de crédito y fecha",
                         ("expr", f"ROWS(Recibidos[{coc['rec']['id']}])"), "#,##0", ">0",
                         '="Los "&TEXT({m},"#,##0")&" pagos recibidos se cruzan por ID y mes '
                         'en cada consulta de cosecha."'))
        big = [f"ROWS({t}[{coc[k]['id']}])" for k, t in (("rec", "Recibidos"), ("esp", "Esperados"))
               if has[k]]
        if big:
            opts.append((A2, "Particionar las tablas de pagos por mes",
                         ("expr", f"MAX({','.join(big)})/{{s_rows}}"), "0.00%", ">=part",
                         '="La tabla de pagos más grande usa "&TEXT({m},"0.00%")&" del límite '
                         'de filas de Excel."'))

    # ---- layout: every block's rows are fixed before anything is written ----
    K0 = 7                                   # first KPI row
    KPI_N = 6
    V = K0 + KPI_N                           # verdict row
    C0 = V + 2; CH = C0 + 3; C1 = CH + 1; Cn = C1 + len(ctrls) - 1
    O0 = Cn + 2; OH = O0 + 3; O1 = OH + 1; On = O1 + len(opts) - 1
    P0 = On + 2; PH = P0 + 3; P1 = PH + 1; Pn = P1 + len(prof) - 1
    S0 = Pn + 2; SH = S0 + 3; S1 = SH + 1; Sn = S1 + len(SUPUESTOS) - 1
    D0 = Sn + 2; DH = D0 + 3; D1 = DH + 1; Dn = D1 + len(LADDER) - 1
    sup = {k: f"${L(5)}${S1 + i}" for i, (k, *_r) in enumerate(SUPUESTOS)}
    pr = lambda c: f"${L(c)}${P1}:${L(c)}${Pn}"
    cr = lambda c: f"${L(c)}${C1}:${L(c)}${Cn}"
    # profile rows are addressed directly: a SUMIFS lookup by column name would read
    # '*', '?' or '~' in a header as wildcards
    prow = {(t, c): P1 + i for i, (t, c, _ty) in enumerate(prof)}
    pc = lambda col, tbl, name: f"${L(col)}${prow[(tbl, name)]}"

    def metric(spec):
        k = spec[0]
        if k == "empty":
            return f'=COUNTIFS({pr(1)},"Data",{pr(7)},0)'
        if k == "prec":
            return f'=COUNTIFS({pr(13)},">"&{sup["prec"]})'
        if k == "intover":
            return f'=COUNTIFS({pr(3)},"entero",{pr(16)},">1")'
        if k == "textover":
            return f'=COUNTIFS({pr(3)},"texto",{pr(16)},">1")'
        if k == "near":
            return f'=COUNTIFS({pr(16)},">="&{sup["alert"]},{pr(16)},"<=1")'
        if k == "iddup":
            return (f'=COUNTIFS({pr(3)},"entero",{pr(11)},">="&{sup["idcard"]},'
                    f'{pr(11)},"<1")')
        if k == "dupid":
            return f"={pc(7, 'Data', spec[1])}-{pc(10, 'Data', spec[1])}"
        if k == "lowuse":
            return f'=COUNTIFS({pr(3)},"texto",{pr(16)},"<"&{sup["lowuse"]})'
        if k == "catalog":
            return f'=COUNTIFS({pr(3)},"texto",{pr(11)},"<"&{sup["card"]},{pr(10)},">0")'
        if k == "idlike":
            return f'=COUNTIFS({pr(3)},"entero",{pr(11)},">="&{sup["idcard"]})'
        if k == "critok":
            return f'=COUNTIFS({pr(4)},"Sí",{pr(8)},0)'
        if k == "expr":
            return "=" + spec[1].replace("{s_rows}", sup["rows"])
        if k == "ratio":
            return f"={spec[1]}/{sup['rows']}"
        if k == "nul":
            return f"={pc(8, spec[1], spec[2])}"
        if k == "nul2":                  # rows missing either field, each row once
            t, a, b = spec[1], spec[2], spec[3]
            return (f"=SUMPRODUCT(--(((LEN(TRIM({t}[{a}]))=0)+"
                    f"(LEN(TRIM({t}[{b}]))=0))>0))")
        if k == "keytype":
            base = pc(3, "Data", coc["credit_id"])
            others = [pc(3, t, coc[s]["id"]) for s, t in (("esp", "Esperados"), ("rec", "Recibidos"))
                      if has[s]]
            return "=--OR(" + ",".join(f"{o}<>{base}" for o in others) + ")"
        raise ValueError(k)

    # ---- 1. Diagnóstico general ----
    block_title(ws, 4, "Diagnóstico general",
                "Completitud, capacidad e integridad en una línea; el detalle sigue abajo.")
    kpis = [
        ("Registros en la cinta", f"=ROWS(Data[{first_col}])", "#,##0"),
        ("Columnas en la cinta", f'=COUNTIF({pr(1)},"Data")', "#,##0"),
        ("Completitud de la cinta",
         f'=IFERROR(SUMIF({pr(1)},"Data",{pr(7)})/SUMIF({pr(1)},"Data",{pr(6)}),0)', "0.00%"),
        ("Campos críticos con nulos", f'=COUNTIFS({pr(4)},"Sí",{pr(8)},">0")', "#,##0"),
        ("Controles en riesgo", f'=COUNTIF({cr(5)},"RIESGO")', "#,##0"),
        ("Controles por revisar", f'=COUNTIF({cr(5)},"REVISAR")', "#,##0"),
    ]
    for i, (lab, f, fmt) in enumerate(kpis):
        put(K0 + i, 1, lab)
        put(K0 + i, 5, f, fmt, bold=True)
    k = lambda i: f"{L(5)}{K0 + i}"
    put(V, 1, (f'="La cinta tiene "&TEXT({k(0)},"#,##0")&" registros en "&{k(1)}&" columnas, '
               f'con completitud de "&TEXT({k(2)},"0.00%")&". "&IF({k(4)}=0,"Ningún control '
               f'está en riesgo","Hay "&{k(4)}&" controles en riesgo")&IF({k(5)}=0,".",'
               f'" y "&{k(5)}&" por revisar.")'), bold=True)

    # ---- 2. Riesgos inmediatos ----
    block_title(ws, C0, "Riesgos Inmediatos",
                "Estado RIESGO o REVISAR cuando el resultado supera el umbral.")
    header_row(ws, CH, ["Área", "Control", "Resultado", "Umbral", "Estado", "Detalle"])
    for i, (area, lab, res, fmt, thr, sev, det) in enumerate(ctrls):
        r = C1 + i
        put(r, 1, area)
        put(r, 2, lab)
        if isinstance(res, tuple) and res[0] == "val":
            put(r, 3, res[1], fmt, blue=True)
        elif isinstance(res, tuple) and res[0] == "arr":
            put(r, 3, res, fmt)
        else:
            put(r, 3, res if isinstance(res, str) else metric(res), fmt)
        put(r, 4, f"={sup[thr]}" if thr else 0, fmt)
        put(r, 5, f'=IF({L(3)}{r}>{L(4)}{r},"{sev}","OK")', bold=True)
        put(r, 6, det)
    fills = {"OK": ("FFE8F3EE", None), "REVISAR": ("FF" + SAND, None),
             "RIESGO": ("FF" + MID, "FFFFFFFF")}
    for word, (bg, fg) in fills.items():
        ws.conditional_formatting.add(cr(5), CellIsRule(
            operator="equal", formula=[f'"{word}"'],
            fill=PatternFill("solid", fgColor=bg),
            font=Font(name="Helvetica", size=8, bold=True, color=fg)))

    # ---- 3. Optimizaciones de rendimiento ----
    block_title(ws, O0, "Optimizaciones de Rendimiento",
                "Mejoras para escalar la cinta a una base de datos; Aplica cuando la métrica lo justifica.")
    header_row(ws, OH, ["Área", "Recomendación", "Aplica", "Métrica", "Evidencia"])
    ntext = f'COUNTIF({pr(3)},"texto")'
    ncrit = f'COUNTIF({pr(4)},"Sí")'
    for i, (area, lab, spec, fmt, when, ev) in enumerate(opts):
        r = O1 + i
        m = f"{L(4)}{r}"
        put(r, 1, area)
        put(r, 2, lab)
        cond = f'{m}>={sup["part"]}' if when == ">=part" else f"{m}{when}"
        put(r, 3, f'=IF({cond},"Sí","No")', bold=True)
        put(r, 4, metric(spec), fmt)
        put(r, 5, ev.format(m=m, ntext=ntext, ncrit=ncrit, s_lowuse=sup["lowuse"]))
    ws.conditional_formatting.add(f"{L(3)}{O1}:{L(3)}{On}", CellIsRule(
        operator="equal", formula=['"Sí"'], fill=PatternFill("solid", fgColor="FF" + PALE)))

    # ---- 4. Perfil por columna ----
    block_title(ws, P0, "Perfil por columna",
                "Nulos, llenado, cardinalidad y capacidad de cada columna, con el tipo y la "
                "restricción que debería declarar una base destino. Distintos no distingue "
                "mayúsculas, igual que Excel.")
    header_row(ws, PH, ["Tabla", "Columna", "Tipo", "Crítico", "Restricción sugerida",
                        "Registros", "Llenos", "Nulos", "% Nulos", "Distintos", "% Distintos",
                        "Long. máx", "Dígitos máx", "Máximo", "Límite por defecto",
                        "Uso del límite", "Tipo recomendado", "Diagnóstico"])
    lad_min = f"${L(1)}${D1}:${L(1)}${Dn}"
    lad_typ = f"${L(2)}${D1}:${L(2)}${Dn}"
    for i, (tbl, col, typ) in enumerate(prof):
        r = P1 + i
        ref = f"{tbl}[{col}]"
        c = lambda n: f"{L(n)}{r}"
        put(r, 1, tbl)
        put(r, 2, col)
        put(r, 3, TYPE_ES.get(typ, typ))
        put(r, 4, "Sí" if (tbl, col) in crit else "No")
        put(r, 5, f'=IF({c(4)}="Sí","NOT NULL (obligatorio)",IF({c(8)}=0,"NOT NULL (candidato)",'
                  f'"NULL permitido"))')
        put(r, 6, f"=ROWS({ref})", "#,##0")
        # TRIM: the load already treats whitespace-only text as empty, so do the same here
        put(r, 7, f"=SUMPRODUCT(--(LEN(TRIM({ref}))>0))", "#,##0")
        put(r, 8, f"={c(6)}-{c(7)}", "#,##0")
        put(r, 9, f"=IFERROR({c(8)}/{c(6)},0)", "0.00%")
        # FILTER drops blanks first: UNIQUE alone returns a blank as 0 and would merge it
        # with a real zero. IF, not IFERROR, guards the empty column so an Excel without
        # UNIQUE shows #NAME? instead of a plausible 0.
        put(r, 10, ("arr", f"=IF({c(7)}=0,0,ROWS(_xlfn.UNIQUE(_xlfn._xlws.FILTER("
                           f"{ref},LEN(TRIM({ref}))>0))))"), "#,##0")
        put(r, 11, f"=IFERROR({c(10)}/{c(7)},0)", "0.00%")
        if typ == "string":
            put(r, 12, f"=SUMPRODUCT(MAX(LEN({ref})))", "0")
            put(r, 15, f"={sup['varchar']}", "#,##0")
            put(r, 16, f"=IFERROR({c(12)}/{c(15)},0)", "0.00%")
            put(r, 17, f"=LOOKUP({c(12)},{lad_min},{lad_typ})")
        elif typ in ("integer", "float"):
            put(r, 14, f"=SUMPRODUCT(MAX(ABS({ref})))",
                "#,##0" if typ == "integer" else "#,##0.00")
            put(r, 13, f'=IFERROR(LEN(TEXT(INT({c(14)}),"0")),0)', "0")
            if typ == "integer":
                put(r, 15, f"={sup['int']}", "#,##0")
                put(r, 16, f"=IFERROR({c(14)}/{c(15)},0)", "0.00%")
                put(r, 17, f'=IF({c(14)}<={sup["smallint"]},"SMALLINT",'
                           f'IF({c(14)}<={sup["int"]},"INT","BIGINT"))')
            else:
                put(r, 15, f"={sup['decimal']}", "0")
                put(r, 16, f"=IFERROR({c(13)}/{c(15)},0)", "0.00%")
                # a float that never leaves [-1, 1] is a rate, same rule as summary_format
                put(r, 17, f'=IF({c(14)}<=1,"DECIMAL(9,6)","DECIMAL("&MIN(38,{c(13)}+2)&",2)")')
        elif typ == "date":
            put(r, 14, f"=MAX({ref})", "dd/mm/yyyy")
            put(r, 17, "DATE")
        else:
            put(r, 17, "BOOLEAN")
        put(r, 18, (f'=IF(AND({c(4)}="Sí",{c(8)}>0),"Nulos en campo crítico",'
                    f'IF({c(7)}=0,"Columna vacía",'
                    f'IF(N({c(13)})>{sup["prec"]},"Excede la precisión de Excel",'
                    f'IF(AND({c(3)}="entero",N({c(16)})>1),"Excede INT: usar BIGINT",'
                    f'IF(AND({c(3)}="texto",N({c(16)})>1),"Excede VARCHAR(255): usar TEXT",'
                    f'IF(N({c(16)})>={sup["alert"]},"Cerca de su límite de tipo",'
                    f'IF(AND({c(3)}="texto",{c(11)}<{sup["card"]},{c(10)}>0),"Candidato a catálogo",'
                    f'IF(AND({c(3)}="texto",N({c(16)})<{sup["lowuse"]}),'
                    f'"Sobreaprovisionado en VARCHAR(255)","Sin observaciones"))))))))'))
    scale(ws, pr(9))
    scale(ws, pr(16))
    databar(ws, pr(7))
    collapse(ws, PH, Pn)

    # ---- 5. Supuestos ----
    block_title(ws, S0, "Supuestos", "Umbrales y límites contra los que se evalúa; editables.")
    header_row(ws, SH, ["Supuesto", "", "", "", "Valor"])
    for i, (_k, lab, val, fmt) in enumerate(SUPUESTOS):
        put(S1 + i, 1, lab)
        put(S1 + i, 5, val, fmt, blue=True)
    collapse(ws, SH, Sn)

    # ---- 6. Escala de tipos de texto ----
    block_title(ws, D0, "Escala de tipos de texto",
                "Longitud mínima a partir de la cual una columna de texto sube de tipo.")
    header_row(ws, DH, ["Longitud desde", "Tipo"])
    for i, (lo, name) in enumerate(LADDER):
        put(D1 + i, 1, lo, "0", blue=True)
        put(D1 + i, 2, name, blue=True)
    collapse(ws, DH, Dn)
    return ws


# ---------- workbook ----------

BANDS = [("<=0", None), (">=1", "<=15"), (">=16", "<=30"), (">=31", "<=60"),
         (">=61", "<=90"), (">=91", "<=120"), (">120", None)]


def criteria(dpd, lo, hi):
    """SUMIFS/COUNTIFS criteria tail for one bucket."""
    return f'"{lo}"' if hi is None else f'"{lo}",Data[{dpd}],"{hi}"'


def bucket_formulas(saldo, dpd, i):
    """(OPB, count) formulas for bucket i.

    A blank days-past-due means the loan is current, so the first bucket also picks up
    empty cells. Without this they match no bucket at all and the panel silently covers
    only the delinquent slice of the book.
    """
    lo, hi = BANDS[i]
    crit = criteria(dpd, lo, hi)
    opb = f"SUMIFS(Data[{saldo}],Data[{dpd}],{crit})"
    cnt = f"COUNTIFS(Data[{dpd}],{crit})"
    if i == 0:
        opb += f'+SUMIFS(Data[{saldo}],Data[{dpd}],"")'
        cnt += f'+COUNTIFS(Data[{dpd}],"")'
    return f"={opb}", f"={cnt}"


def clear(ws, ref):
    for row in ws[ref]:
        for c in row:
            c.value = None
            c.style = "Normal"


def all_columns():
    return HEADERS + [n for n, _ in CALCS]


def summary_format(col):
    """Number format for a summary row.

    A float that never leaves [-1, 1] is a rate, not money: 'Tasa' 0.355 reads as
    35.50%, not $0.36. Anything wider gets the accounting format.
    """
    t = TYPES[col]
    if t == "date":
        return MMMYY
    if t in ("integer", "boolean"):
        return "0"
    if t != "float":
        return "General"
    lo, hi = DB.execute(f"SELECT MIN({q(col)}), MAX({q(col)}) FROM tape").fetchone()
    return "0.00%" if lo is not None and lo >= -1 and hi <= 1 else ACCT


def build_workbook(saldo, dpd, summary, par=(), hist=(), coc=None, conc=()):
    wb = load_workbook(TEMPLATE)
    wb.loaded_theme = theme_xml(getattr(wb, "loaded_theme", None))
    ws = wb["Portfolio"]
    # the template already starts its content at D, so the three spacer columns
    # just get the standard width here rather than a new layout
    for i in range(1, OFF + 1):
        ws.column_dimensions[get_column_letter(i)].width = MARGIN_W
    cols = all_columns()
    used = {s.lower() for s in wb.sheetnames}

    # --- data sheet as a real Excel Table ---
    ds = wb.create_sheet("Data")
    margins(ds)
    ds.append([None] * OFF + cols)
    fmt = {"date": "dd/mm/yyyy", "float": "#,##0.00", "integer": "#,##0"}
    for r in DB.execute(f"SELECT {','.join(q(c) for c in cols)} FROM tape"):
        out = []
        for c, v in zip(cols, r):
            if TYPES[c] == "date" and v:
                out.append(date.fromisoformat(v))
            elif TYPES[c] == "boolean" and v is not None:
                out.append(bool(v))
            else:
                out.append(v)
        ds.append([None] * OFF + out)
    last = f"{L(len(cols))}{ds.max_row}"
    for i, c in enumerate(cols, start=1):
        nf = fmt.get(TYPES[c])
        if nf:
            for cell in ds.iter_rows(min_row=2, min_col=(i) + OFF, max_col=(i) + OFF):
                cell[0].number_format = nf
        ds.column_dimensions[L(i)].width = max(10, min(28, len(c) + 4))
    t = Table(displayName="Data", ref=f"{L(1)}1:{last}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ds.add_table(t)
    ds.freeze_panes = f"{L(1)}2"

    # --- buckets: the only cells the template left for a human to paste into ---
    for i in range(len(BANDS)):
        r = 15 + i
        ws[f"J{r}"], ws[f"L{r}"] = bucket_formulas(saldo, dpd, i)

    # --- summary: fixed labels replaced by the selected columns ---
    ws["G13"] = "Wtd. Average"      # it is no longer a simple mean; say so
    for i in range(6):
        r = 14 + i
        if i < len(summary):
            c = summary[i]
            ws[f"D{r}"] = c
            for col, fn in (("E", "MIN"), ("F", "MAX"), ("G", None)):
                cell = ws[f"{col}{r}"]
                if r == 19:                     # row 19 has no E:G in the template
                    cell._style = copy(ws[f"{col}15"]._style)
                # the average is weighted by outstanding balance: a simple mean lets a
                # MXN $10k loan move the portfolio rate as much as a MXN $2mm one
                cell.value = (f"={fn}(Data[{c}])" if fn else
                              f"=IFERROR(SUMPRODUCT(Data[{c}],Data[{saldo}])"
                              f"/SUM(Data[{saldo}]),0)")
                cell.number_format = summary_format(c)
        elif r != 19:                       # row 19 has no E:G to clear
            ws[f"D{r}"] = None
            for col in "EFG":
                ws[f"{col}{r}"].value = None
        else:
            ws["D19"] = None

    # --- totals block: only the two rows that survive without a default definition ---
    ws["D23"] = "Active Total Credits"
    ws["E23"] = f'=COUNTIFS(Data[{saldo}],">0")'
    ws["F23"] = f'=SUMIFS(Data[{saldo}],Data[{saldo}],">0")'
    ws["G23"] = "=IFERROR(F23/$F$24,0)"
    ws["D24"] = "Outstanding Balance"
    ws["E24"] = f"=COUNTA(Data[{saldo}])"
    ws["F24"] = f"=SUM(Data[{saldo}])"
    ws["G24"] = "=IFERROR(F24/$F$24,0)"
    clear(ws, "D25:G27")

    # --- borrowing base removed with the advance rates ---
    clear(ws, "O11:Q21")

    # openpyxl keeps the template's drawing; only re-add if a version drops it
    if not ws._images:
        try:
            with zipfile.ZipFile(TEMPLATE) as z:
                img = XLImage(io.BytesIO(z.read("xl/media/image1.png")))
            img.width, img.height = 104, 111
            ws.add_image(img, "B2")
        except Exception:
            pass  # ponytail: logo is cosmetic, never fail a generate over it

    used.add("data")
    stats = {}
    if coc:
        cmap = cohort_map(coc["credit_id"], coc["orig_date"])
        _, esp_n = add_pay_sheet(wb, used, "esperados", "Esperados", coc["esp"],
                                 coc["credit_id"], dedupe=coc["dedupe"])
        _, rec_n = add_pay_sheet(wb, used, "recibidos", "Recibidos", coc["rec"],
                                 coc["credit_id"], dedupe=False, cmap=cmap)
        _, credits_n = add_credits_sheet(wb, used, coc["credit_id"], saldo, coc["orig"],
                                         coc["orig_date"], coc["rec"])
        add_moic_sheet(wb, used, coc["orig_date"])
        add_coc_sheet(wb, used, coc["esp"], coc["rec"], esp_n, rec_n, coc["orig_date"])
        stats = {"credits": credits_n, "esperados_rows": esp_n, "recibidos_rows": rec_n,
                 "esperados_dropped": PAY["esperados"]["rows"] - esp_n,
                 "recibidos_dropped": PAY["recibidos"]["rows"] - rec_n,
                 "detected": {k: {x: coc[k][x] for x in ("id", "date", "amount")}
                              for k in ("esp", "rec")}}
    pst = {}
    if coc:
        # facts that only exist at load time: the dropped rows never reach the workbook.
        # Everything is counted in raw rows and inside the tape's scope, so duplicates of
        # orphan rows are not also reported as duplicates of the tape's own credits.
        in_tape = tape_ids(coc["credit_id"])

        def scope(slot, key):
            idc = q(coc[key]["id"])
            null = DB.execute(f"SELECT COUNT(*) FROM {slot} WHERE {idc} IS NULL "
                              f"OR TRIM(CAST({idc} AS TEXT)) = ''").fetchone()[0]
            inside = DB.execute(f"SELECT COUNT(*) FROM {slot} "
                                f"WHERE CAST({idc} AS TEXT) IN {in_tape}").fetchone()[0]
            return PAY[slot]["rows"], null, inside

        e_raw, e_null, e_in = scope("esperados", "esp")
        r_raw, r_null, r_in = scope("recibidos", "rec")
        e_in_distinct = DB.execute(
            f"SELECT COUNT(*) FROM (SELECT DISTINCT * FROM esperados WHERE "
            f"CAST({q(coc['esp']['id'])} AS TEXT) IN {in_tape})").fetchone()[0]
        pst = {"esp_n": esp_n, "rec_n": rec_n, "credits": credits_n,
               "esp_dups": e_in - e_in_distinct,
               "esp_null_ids": e_null, "rec_null_ids": r_null,
               "esp_orphans": e_raw - e_in - e_null,
               "rec_orphans": r_raw - r_in - r_null}
    if par:
        add_par_sheets(wb, used, par, saldo, dpd)
    if conc:
        add_concentration_sheet(wb, used, conc, saldo)
    if hist:
        add_histogram_sheet(wb, used, hist)
    add_audit_sheet(wb, used, saldo, dpd, coc, pst)

    # analysis first, raw reference tables last
    ref = ("Credits", "Data", "Esperados", "Recibidos")
    wb._sheets = ([x for x in wb._sheets if x.title not in ref]
                  + [x for x in wb._sheets if x.title in ref])

    wb.calculation.fullCalcOnLoad = True
    path = os.path.join(tempfile.gettempdir(), f"portfolio_{uuid.uuid4().hex}.xlsx")
    wb.save(path)
    return path, stats


# ---------- check ----------

def selftest():
    assert parse_number("$ 1,234.56") == 1234.56
    assert parse_number("1.234,56") == 1234.56
    assert parse_number("1,234") == 1234
    assert parse_number("1,5") == 1.5
    # a lone dot is a decimal point, even with exactly three digits after it
    assert parse_number("0.425") == 0.425
    assert parse_number("501129.306") == 501129.306
    assert parse_number("1.234.567") == 1234567
    assert parse_number("(1,234.56)") == -1234.56
    assert parse_number("MXN $12,000") == 12000
    assert parse_number("45%") == 0.45
    assert parse_number("abc") is None
    assert parse_date("18/09/2026") == date(2026, 9, 18)
    assert parse_date("2026-09-18") == date(2026, 9, 18)
    assert parse_date("13/13/2026") is None
    assert parse_int("1,234") == 1234 and parse_int("1.5") is None
    assert detect_type(["1", "2", "3"]) == "integer"
    assert detect_type(["1.5", "2"]) == "float"
    assert detect_type(["18/09/2026", ""]) == "date"
    assert detect_type(["si", "no"]) == "boolean"
    assert detect_type(["a", "1"]) == "string"
    assert criteria("dpd", "<=0", None) == '"<=0"'
    assert criteria("dpd", ">=1", "<=15") == '">=1",Data[dpd],"<=15"'
    assert clean_name("saldo [mxn]", set()) == "saldo mxn"
    # the seven buckets must partition every integer, so nothing is dropped
    def bucket(d):
        hits = [i for i, (lo, hi) in enumerate(BANDS)
                if eval(f"{d}{lo}") and (hi is None or eval(f"{d}{hi}"))]
        return hits
    for d in range(-5, 200):
        assert len(bucket(d)) == 1, (d, bucket(d))
    # a blank credit id is a missing id: it must not join, and must not be counted as
    # both null and matched (tape A, B, '' against payments A, A, B, '', '', NULL, Z)
    global DB
    saved, DB = DB, sqlite3.connect(":memory:")
    DB.execute('CREATE TABLE tape ("cid" TEXT)')
    DB.executemany("INSERT INTO tape VALUES (?)", [("A",), ("B",), ("",)])
    DB.execute('CREATE TABLE p ("id" TEXT)')
    DB.executemany("INSERT INTO p VALUES (?)", [(x,) for x in ("A", "A", "B", "", "", None, "Z")])
    inside = DB.execute(f'SELECT COUNT(*) FROM p WHERE CAST("id" AS TEXT) IN {tape_ids("cid")}').fetchone()[0]
    null = DB.execute('SELECT COUNT(*) FROM p WHERE "id" IS NULL OR TRIM("id") = \'\'').fetchone()[0]
    assert (inside, null, 7 - inside - null) == (3, 3, 1), (inside, null)
    DB = saved
    print("ok")


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        selftest()
    elif "--dev" in sys.argv:
        app.run(debug=True)             # auto-reload and the Werkzeug debugger, local dev only
    else:
        # the shipped default: no debugger, bound to 127.0.0.1 so nothing leaves the
        # machine, and the browser opens itself once the server is listening
        import threading, webbrowser
        threading.Timer(1.5, lambda: webbrowser.open("http://127.0.0.1:5000")).start()
        app.run(debug=False)
