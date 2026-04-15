import io, os, re, uuid
from datetime import date
from flask import Flask, request, redirect, url_for, render_template_string, session, send_file, abort
import pandas as pd

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-" + uuid.uuid4().hex)
STORAGE = {}

# ---------------- Name helpers ----------------
ACRONYM_CANON = {
    "edp": "EDP", "edt": "EDT", "edc": "EDC", "ysl": "YSL", "ck": "CK", "dkny": "DKNY",
    "spf": "SPF", "uv": "UV", "uva": "UVA", "uvb": "UVB", "egf": "EGF", "sd": "SD",
    "hdmi": "HDMI", "usb": "USB", "ph": "pH"
}
SMALL_WORDS = {"de", "des", "du", "le", "la", "les", "pour", "aux", "and", "of", "the", "for", "with", "von", "di", "da", "del", "in", "on", "at", "by", "to"}

def _title_token(tok: str, first_in_seg: bool) -> str:
    if not tok:
        return tok
    low = tok.lower()
    if low in ACRONYM_CANON:
        return ACRONYM_CANON[low]
    if not first_in_seg and low in SMALL_WORDS:
        return low
    return tok[0].upper() + tok[1:].lower()

def _every_word_cap(tok: str) -> str:
    if not tok:
        return tok
    low = tok.lower()
    if low in ACRONYM_CANON:
        return ACRONYM_CANON[low]
    return tok[0].upper() + tok[1:].lower()

def normalize_name(text, title_every_word=True):
    if text is None:
        return "", ""

    s = str(text).strip()
    if not s:
        return "", ""

    s = re.sub(r"\s+", " ", s)
    s = s.replace("’", "'").replace("`", "'")
    s = s.replace("—", "-").replace("–", "-").replace("•", "-").replace("|", "-").replace("/", "-")
    s = s.replace('"', "").replace("“", "").replace("”", "")

    s = re.sub(r"\s*-\s*", " – ", s)
    s = re.sub(r"\bNew\b|\b100% Original\b|\bFree Shipping\b|#[A-Za-z0-9_]+|[\U0001F300-\U0001FAFF]", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+", " ", s).strip()

    # маха "0 Ml", "0 ML", "0ml", "0.0 Ml" и подобни
    s = re.sub(r"\b0(?:[.,]0+)?\s*ML\b", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\b0(?:[.,]0+)?\s*OZ\b", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+", " ", s).strip()

    # нормализира количества
    s = re.sub(r"\b([0-9]+)\s*ML\b", r"\1 ml", s, flags=re.IGNORECASE)
    s = re.sub(r"\b([0-9.]+)\s*OZ\b", r"\1 oz", s, flags=re.IGNORECASE)

    segs = [p.strip() for p in s.split(" – ")] if " – " in s else [s]
    out = []
    for seg in segs:
        words = []
        for w in seg.split(" "):
            if "-" in w:
                parts = w.split("-")
                parts = [
                    _every_word_cap(p) if title_every_word else _title_token(p, first_in_seg=(len(words) == 0 and i == 0))
                    for i, p in enumerate(parts)
                ]
                words.append("-".join(parts))
            else:
                words.append(_every_word_cap(w) if title_every_word else _title_token(w, first_in_seg=(len(words) == 0)))
        out.append(" ".join(words))

    s = " – ".join(out)
    s = s.replace("Eau De Parfum", "Eau de Parfum").replace("Eau De Toilette", "Eau de Toilette").replace("Eau De Cologne", "Eau de Cologne")
    s = re.sub(r"\(\s*\)", "", s)
    s = re.sub(r"\[\s*\]", "", s)
    s = re.sub(r"\s+[-–]\s*$", "", s)
    s = re.sub(r"^\s+[-–]\s+", "", s)
    s = re.sub(r"\s+", " ", s).strip()

    full = s
    if len(s) > 60:
        s = s[:60] + "…"
    return s, full

# ---------------- Numbers / Price / EAN ----------------
def clean_price(v):
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    s = s.replace("€", "").replace("$", "").replace("£", "")
    s = re.sub(r"[^\d,.\-]", "", s).replace(" ", "").replace(",", ".")
    try:
        x = float(s)
    except ValueError:
        return None
    return round(x + 1e-7, 2)

def fmt_price(val, currency="EUR", primary=False, strike=False):
    if val is None:
        return ""
    symbol_map = {
        "NONE": "",
        "EUR": "€",
        "USD": "$",
    }
    symbol = symbol_map.get(currency, "€")
    num = f"{val:.2f}{symbol}"
    if primary:
        return f'<span style="color:#dc2626;font-size:13px;font-weight:700">{num}</span>'
    if strike:
        return f'<span style="text-decoration:line-through;color:#6b7280">{num}</span>'
    return num

def ean13_normalize(s):
    if s is None:
        return ""
    d = re.sub(r"\D", "", str(s))
    if not d:
        return ""
    if len(d) < 13:
        d = d.zfill(13)
    elif len(d) > 13:
        return ""
    return d

# ---------------- HTML builder ----------------
def build_html(df, plan, currency="EUR"):
    table_style = "max-width:600px;margin:0 auto;border-collapse:collapse;font-family:Arial,Helvetica,sans-serif;font-size:11px;line-height:1.2;table-layout:auto;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%"
    th_style = "border:1px solid #e5e7eb;padding:4px 6px;background:#f8fafc;vertical-align:middle;text-align:left;white-space:nowrap"
    td_text = "border:1px solid #e5e7eb;padding:4px 6px;vertical-align:middle;text-align:left;white-space:nowrap;overflow:hidden;text-overflow:ellipsis"
    td_num = td_text + ";text-align:right;width:1%"

    used = [p for p in plan if p["use"]]

    # Коя колона да се нормализира като име:
    # 1) ако има роля Name -> тя
    # 2) ако няма, а има EAN -> следващата използвана след EAN
    # 3) иначе втората използвана
    name_srcs = {p["src"] for p in used if p["role"] == "Name"}
    if not name_srcs:
        ean_index = next((i for i, p in enumerate(used) if p["role"] == "EAN"), None)
        if ean_index is not None and ean_index + 1 < len(used):
            name_srcs = {used[ean_index + 1]["src"]}
        elif len(used) >= 2:
            name_srcs = {used[1]["src"]}

    role_src = {p["role"]: p["src"] for p in used if p["role"] in {"Price", "Promo Price"}}

    rows = []
    for _, row in df.iterrows():
        price_val = clean_price(row[role_src["Price"]]) if "Price" in role_src else None
        promo_val = clean_price(row[role_src["Promo Price"]]) if "Promo Price" in role_src else None

        if promo_val is not None and promo_val > 0:
            promo_html = fmt_price(promo_val, currency=currency, primary=True)
            price_html = fmt_price(price_val, currency=currency, strike=True) if price_val is not None else ""
        else:
            promo_html = ""
            price_html = fmt_price(price_val, currency=currency)

        cells = []
        empty = True

        for p in plan:
            if not p["use"]:
                continue

            src, role = p["src"], p["role"]
            val = row[src] if src in df.columns else ""

            if src in name_srcs:
                short, full = normalize_name(val, title_every_word=True)
                cells.append(f'<td style="{td_text};max-width:360px;" title="{full}">{short}</td>')
                if short:
                    empty = False
                continue

            if role in {"Stock", "QTY"}:
                s = re.sub(r"[^\d]", "", str(val))
                cells.append(f'<td style="{td_num}">{s}</td>' if s else f'<td style="{td_num}"></td>')
                if s:
                    empty = False

            elif role == "EAN":
                s = ean13_normalize(val)
                cells.append(f'<td style="{td_text}">{s}</td>')
                if s:
                    empty = False

            elif role == "Price":
                cells.append(f'<td style="{td_num}">{price_html}</td>' if price_html else f'<td style="{td_num}"></td>')
                if price_html:
                    empty = False

            elif role == "Promo Price":
                cells.append(f'<td style="{td_num}">{promo_html}</td>' if promo_html else f'<td style="{td_num}"></td>')
                if promo_html:
                    empty = False

            else:
                v = str(val).strip()
                cells.append(f'<td style="{td_text}">{v}</td>' if v else f'<td style="{td_text}"></td>')
                if v:
                    empty = False

        if empty:
            continue

        rows.append("<tr>" + "".join(cells) + "</tr>")

    def th_align(role):
        return th_style if role not in {"Stock", "QTY", "Price", "Promo Price"} else th_style.replace("text-align:left", "text-align:right")

    thead = "".join([f'<th style="{th_align(p["role"])}">{p["header"]}</th>' for p in plan if p["use"]])
    tbody = "".join(rows)
    html = f'<table role="presentation" style="{table_style}"><thead><tr>{thead}</tr></thead><tbody>{tbody}</tbody></table>'
    return re.sub(r">\s+<", "><", html)

# ---------------- UI (grey + pink) ----------------
def wrap(content_html: str) -> str:
    T_BASE = """
<!doctype html><html><head><meta charset="utf-8"><title>Excel → HTML Table</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet">
<style>
:root{--bg:#f5f6f8;--panel:#fff;--muted:#6b7280;--text:#0f172a;--border:#e5e7eb;--pink:#ec4899;--pink-hover:#f472b6;--pink-soft:#fde7f3}
*{box-sizing:border-box} body{font-family:Inter,Arial,Helvetica,sans-serif;background:var(--bg);color:var(--text);margin:0}
.header{position:sticky;top:0;background:#ffffffcc;backdrop-filter:blur(8px);border-bottom:1px solid var(--border);z-index:10}
.header .wrap{max-width:1100px;margin:0 auto;padding:14px 20px;display:flex;gap:12px;align-items:center}
.brand{font-weight:700;letter-spacing:.2px}.pill{background:var(--pink-soft);color:#be185d;padding:4px 10px;border-radius:999px;font-size:12px}
.container{max-width:1100px;margin:24px auto;padding:0 20px}
.card{background:var(--panel);border:1px solid var(--border);border-radius:14px;padding:18px;box-shadow:0 6px 18px rgba(0,0,0,.05)}
h2{margin:18px 0 10px 0;font-size:22px} label{font-weight:600}
input[type=text],input[type=file],select{width:100%;padding:10px 12px;background:#fff;border:1px solid var(--border);border-radius:10px;color:#0f172a}
input[type=text]:focus,select:focus{outline:none;border-color:var(--pink);box-shadow:0 0 0 3px rgba(236,72,153,.2)}
.btn{display:inline-flex;gap:8px;align-items:center;background:var(--pink);color:#fff;padding:10px 14px;border-radius:10px;text-decoration:none;border:none;cursor:pointer;font-weight:700;letter-spacing:.2px}
.btn:hover{background:var(--pink-hover)} .btn.secondary{background:#e5e7eb;color:#111827}
.small{font-size:12px;color:var(--muted)} .badge{display:inline-block;background:#fff;border:1px solid var(--border);color:#111827;padding:4px 10px;border-radius:999px;font-size:12px;margin:3px 4px}
.cols{width:100%;border-radius:10px;overflow:hidden;border:1px solid var(--border)} .cols th{background:#f8fafc}
th,td{border:1px solid var(--border);padding:8px 10px} .role{width:180px} .use{text-align:center;width:90px}
.row{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.preview{border:1px solid var(--border);border-radius:12px;overflow:auto;max-height:520px;background:#fff}
</style></head><body>
<div class="header"><div class="wrap"><div class="brand">Excel → HTML Table</div><span class="pill">Grey ✕ Pink</span></div></div>
<div class="container">{CONTENT}</div></body></html>
""".strip()
    return T_BASE.replace("{CONTENT}", content_html)

T_UPLOAD = wrap("""
<div class="card">
  <h2>1) Качи Excel (.xlsx)</h2>
  <form action="{{ url_for('upload') }}" method="post" enctype="multipart/form-data">
    <div class="row">
      <div>
        <label>Файл</label><br>
        <input type="file" name="file" accept=".xlsx" required>
      </div>
    </div>
    <p class="small">Файлът се обработва локално.</p>
    <button class="btn">Качи</button>
  </form>
</div>
""")

T_FOUND = wrap("""
<div class="card">
  <h2>2) Намерени листове в Excel файла</h2>
  <p class="small">Файлът съдържа <b>{{ sheet_count }}</b> лист(а). Избери кой лист искаш да обработиш:</p>

  <form action="{{ url_for('select_sheet') }}" method="post">
    <div style="max-width:420px;margin:12px 0">
      <label>Лист</label>
      <select name="sheet_name" required>
        {% for s in sheet_names %}
          <option value="{{ s }}">{{ s }}</option>
        {% endfor %}
      </select>
    </div>

    <button class="btn">Продължи</button>
    <a href="{{ url_for('index') }}" class="btn secondary">Назад</a>
  </form>
</div>
""")

T_PLAN = wrap("""
<div class="card">
  <h2>3) Настрой за всяка колона</h2>
  <form action="{{ url_for('generate') }}" method="post">
    <table class="cols">
      <thead>
        <tr>
          <th>#</th>
          <th>Оригинално име</th>
          <th>Име в крайния файл</th>
          <th class="role">Роля</th>
          <th class="use">Ползвай</th>
        </tr>
      </thead>
      <tbody>
      {% for h in headers %}
        <tr>
          <td>{{ loop.index }}</td>
          <td>{{ h }}<input type="hidden" name="src_{{ loop.index0 }}" value="{{ h }}"></td>
          <td><input type="text" name="hdr_{{ loop.index0 }}" value="{{ h }}"></td>
          <td>
            <select name="role_{{ loop.index0 }}" class="role">
              {% set det = auto_roles.get(h.lower(), 'Text') %}
              {% for r in roles %}
                <option value="{{ r }}" {% if det==r %}selected{% endif %}>{{ r }}</option>
              {% endfor %}
            </select>
          </td>
          <td class="use"><input type="checkbox" name="use_{{ loop.index0 }}" value="1" checked></td>
        </tr>
      {% endfor %}
      </tbody>
    </table>

    <hr>

    <div style="max-width:320px">
      <label>Валута</label>
      <select name="currency">
        <option value="EUR" selected>EUR (€)</option>
        <option value="USD">USD ($)</option>
        <option value="NONE">Без валута</option>
      </select>
    </div>

    <div style="margin-top:14px">
      <button class="btn">Генерирай HTML</button>
      <a href="{{ url_for('index') }}" class="btn secondary">Начало</a>
    </div>
  </form>
</div>
""")

T_RESULT = wrap("""
<div class="card">
  <h2>4) Резултат</h2>
  <p class="small">Обработен лист: <b>{{ sheet_name }}</b></p>
  <p>
    <a class="btn" href="{{ url_for('download_html', sid=sid) }}">Свали products_table_{{ today }}.html</a>
    <a href="{{ url_for('index') }}" class="btn secondary">Ново качване</a>
  </p>
  <p class="small">Преглед:</p>
  <div class="preview">{{ preview|safe }}</div>
</div>
""")

# ---------------- Helpers ----------------
def detect_role(h):
    hl = h.lower().strip()
    if hl in {"ean", "barcode", "ean13"}:
        return "EAN"
    if hl in {"name", "product", "title"}:
        return "Name"
    if hl in {"qty", "quantity"}:
        return "QTY"
    if hl in {"stock", "available", "onhand"}:
        return "Stock"
    if hl in {"price", "regular price", "list price"}:
        return "Price"
    if hl in {"promo", "promo price", "sale price", "discount price"}:
        return "Promo Price"
    return "Text"

# ---------------- Routes ----------------
@app.route("/")
def index():
    session.clear()
    session["sid"] = uuid.uuid4().hex
    return render_template_string(T_UPLOAD)

@app.route("/upload", methods=["POST"])
def upload():
    try:
        f = request.files.get("file")
        print("DEBUG file:", f)
        print("DEBUG filename:", getattr(f, "filename", None))

        if not f or not f.filename.lower().endswith(".xlsx"):
            return "Моля качи .xlsx файл.", 400

        sid = session.get("sid") or uuid.uuid4().hex
        session["sid"] = sid
        print("DEBUG sid:", sid)

        data = f.read()
        print("DEBUG bytes:", len(data))

        STORAGE[sid] = {"xlsx": data}

        excel_file = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
        sheet_names = excel_file.sheet_names
        print("DEBUG sheets:", sheet_names)

        STORAGE[sid]["sheet_names"] = sheet_names

        return render_template_string(
            T_FOUND,
            sheet_names=sheet_names,
            sheet_count=len(sheet_names)
        )

    except Exception as e:
        import traceback
        traceback.print_exc()
        return f"<h1>UPLOAD ERROR</h1><pre>{str(e)}</pre>", 500

@app.route("/select_sheet", methods=["POST"])
def select_sheet():
    sid = session.get("sid")
    if not sid or sid not in STORAGE:
        return redirect(url_for("index"))

    sheet_name = request.form.get("sheet_name")
    if not sheet_name:
        return "Не е избран лист.", 400

    data = STORAGE[sid]["xlsx"]
    df = pd.read_excel(
        io.BytesIO(data),
        sheet_name=sheet_name,
        dtype=str,
        engine="openpyxl"
    ).fillna("")

    STORAGE[sid]["selected_sheet"] = sheet_name
    STORAGE[sid]["df"] = df
    STORAGE[sid]["headers"] = list(df.columns)

    auto_roles = {h.lower(): detect_role(h) for h in STORAGE[sid]["headers"]}
    roles = ["Text", "EAN", "Name", "Stock", "QTY", "Price", "Promo Price"]

    return render_template_string(
        T_PLAN,
        headers=STORAGE[sid]["headers"],
        auto_roles=auto_roles,
        roles=roles
    )

@app.route("/generate", methods=["POST"])
def generate():
    sid = session.get("sid")
    if not sid or sid not in STORAGE:
        return redirect(url_for("index"))

    df = STORAGE[sid]["df"].copy()
    headers = STORAGE[sid]["headers"]

    plan = []
    for i, _ in enumerate(headers):
        src = request.form.get(f"src_{i}")
        hdr = (request.form.get(f"hdr_{i}") or "").strip()
        role = request.form.get(f"role_{i}") or "Text"
        use = (request.form.get(f"use_{i}") == "1")

        if not src:
            continue
        if not hdr:
            hdr = src

        plan.append({
            "src": src,
            "header": hdr,
            "role": role,
            "use": use
        })

    if not any(p["use"] for p in plan):
        return "Не е избрана нито една колона за крайния файл.", 400

    currency = request.form.get("currency", "EUR")
    html = build_html(df, plan, currency=currency)

    STORAGE[sid]["html"] = html
    STORAGE[sid]["filename"] = f"products_table_{date.today().isoformat()}.html"

    return render_template_string(
        T_RESULT,
        sid=sid,
        today=date.today().isoformat(),
        preview=html,
        sheet_name=STORAGE[sid].get("selected_sheet", "")
    )

@app.route("/download/<sid>")
def download_html(sid):
    if sid not in STORAGE or "html" not in STORAGE[sid]:
        abort(404)
    html = STORAGE[sid]["html"].encode("utf-8")
    fn = STORAGE[sid].get("filename", f"products_table_{date.today().isoformat()}.html")
    return send_file(io.BytesIO(html), mimetype="text/html", as_attachment=True, download_name=fn)

if __name__ == "__main__":
    app.run(debug=True)
