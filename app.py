import ast
import base64
import copy
import io
import json
import logging
import math
import os
import re
import sqlite3
import tempfile
import warnings
from datetime import date, datetime
from urllib.parse import urlencode

import openpyxl
from dotenv import load_dotenv
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Border
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.datetime import from_excel, to_excel
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.formula import ArrayFormula
from pycel import ExcelCompiler

from excel_functions import datedif
from flask import Flask, abort, g, jsonify, redirect, render_template, request, url_for
from google import genai

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))  # GEMINI_API_KEY
logging.getLogger("pycel").setLevel(logging.CRITICAL)  # rumus gagal dihitung tampil sebagai #ERR, tidak perlu log

app = Flask(__name__)
app.config["DB"] = os.path.join(os.path.dirname(os.path.abspath(__file__)), "onavi.db")
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # batas upload PDF 20 MB

# Menu sidebar. Tambah menu baru: tambah satu baris di sini + satu route + satu template.
MENU = [
    ("dashboard", "Dashboard", "bi-grid-1x2"),
    ("projects", "Proyek & kontrak", "bi-briefcase"),
    ("sales", "Sales", "bi-graph-up-arrow"),
    ("storage", "Kalkulator storage", "bi-hdd-stack"),
]

STATUSES = ["Belum Mulai", "Berjalan", "Selesai", "Batal"]

SCHEMA = """
CREATE TABLE IF NOT EXISTS projects (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL,
    client TEXT NOT NULL,
    contract_no TEXT,
    location TEXT,
    contract_value REAL NOT NULL DEFAULT 0,
    est_cost REAL NOT NULL DEFAULT 0,
    status TEXT NOT NULL DEFAULT 'Belum Mulai',
    start_date TEXT,
    end_date TEXT
);
CREATE TABLE IF NOT EXISTS storage_logs (
    id INTEGER PRIMARY KEY,
    used_at TEXT NOT NULL,
    query TEXT NOT NULL UNIQUE,
    summary TEXT NOT NULL,
    total REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS sales (
    id INTEGER PRIMARY KEY,
    sale_date TEXT NOT NULL,
    product TEXT NOT NULL,
    location TEXT NOT NULL,
    qty INTEGER NOT NULL DEFAULT 1,
    amount REAL NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS monitoring_projects (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL,
    xlsx_path TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS monitoring_costs (
    id INTEGER PRIMARY KEY,
    project_id INTEGER NOT NULL,
    month TEXT NOT NULL,
    plan REAL NOT NULL DEFAULT 0,
    actual REAL NOT NULL DEFAULT 0,
    cash_in REAL NOT NULL DEFAULT 0,
    UNIQUE (project_id, month)
);
CREATE TABLE IF NOT EXISTS monitoring_books (
    project_id INTEGER PRIMARY KEY,
    xlsx BLOB NOT NULL,
    imported_at TEXT NOT NULL,
    extra_rows TEXT NOT NULL DEFAULT '{}',
    extra_cols TEXT NOT NULL DEFAULT '{}',
    month_axes TEXT NOT NULL DEFAULT '{}'
);
CREATE TABLE IF NOT EXISTS pum_template (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    xlsx BLOB NOT NULL,
    source TEXT NOT NULL,
    imported_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS pum_docs (
    id INTEGER PRIMARY KEY,
    project_id INTEGER NOT NULL,
    xlsx BLOB NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    doc_date TEXT,
    no_pum TEXT,
    periode TEXT,
    total REAL NOT NULL DEFAULT 0,
    categories TEXT NOT NULL DEFAULT '{}'
);
CREATE TABLE IF NOT EXISTS signatures (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL UNIQUE COLLATE NOCASE,
    image BLOB NOT NULL,
    mime TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS pum_controllers (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS monitoring_edits (
    project_id INTEGER NOT NULL,
    sheet TEXT NOT NULL,
    cell TEXT NOT NULL,
    value TEXT,
    PRIMARY KEY (project_id, sheet, cell)
);
"""


def db():
    if "db" not in g:
        g.db = sqlite3.connect(app.config["DB"])
        g.db.row_factory = sqlite3.Row
        g.db.executescript(SCHEMA)
        # database lama (dibuat sebelum ada kolom end_date)
        if "end_date" not in [c["name"] for c in g.db.execute("PRAGMA table_info(projects)")]:
            g.db.execute("ALTER TABLE projects ADD COLUMN end_date TEXT")
        # riwayat kalkulator storage dibuat sebelum ada kolom nama
        if "name" not in [c["name"] for c in g.db.execute("PRAGMA table_info(storage_logs)")]:
            g.db.execute("ALTER TABLE storage_logs ADD COLUMN name TEXT")
        # proyek monitoring dibuat sebelum ada input manual: sumber 'xlsx' | 'manual', info = JSON isian manual
        if "source" not in [c["name"] for c in g.db.execute("PRAGMA table_info(monitoring_projects)")]:
            g.db.execute("ALTER TABLE monitoring_projects ADD COLUMN source TEXT NOT NULL DEFAULT 'xlsx'")
            g.db.execute("ALTER TABLE monitoring_projects ADD COLUMN info TEXT NOT NULL DEFAULT '{}'")
        # PUM dibuat sebelum ada tanda tangan: {slot: nama penandatangan}
        if "signatures" not in [c["name"] for c in g.db.execute("PRAGMA table_info(pum_docs)")]:
            g.db.execute("ALTER TABLE pum_docs ADD COLUMN signatures TEXT NOT NULL DEFAULT '{}'")
        # PUM dibuat sebelum ada nomor otomatis: nomor urut & tahun nomor PUM
        if "seq" not in [c["name"] for c in g.db.execute("PRAGMA table_info(pum_docs)")]:
            g.db.execute("ALTER TABLE pum_docs ADD COLUMN seq INTEGER")
            g.db.execute("ALTER TABLE pum_docs ADD COLUMN year INTEGER")
        # snapshot dibuat sebelum ada tambah/hapus baris: extra_rows = {judul sheet: jumlah baris bertambah}
        if "extra_rows" not in [c["name"] for c in g.db.execute("PRAGMA table_info(monitoring_books)")]:
            g.db.execute("ALTER TABLE monitoring_books ADD COLUMN extra_rows TEXT NOT NULL DEFAULT '{}'")
        # snapshot dibuat sebelum jumlah kolom bulan mengikuti kontrak: extra_cols = {judul sheet: kolom bertambah}
        if "extra_cols" not in [c["name"] for c in g.db.execute("PRAGMA table_info(monitoring_books)")]:
            g.db.execute("ALTER TABLE monitoring_books ADD COLUMN extra_cols TEXT NOT NULL DEFAULT '{}'")
        # posisi tabel bulan disimpan, agar tetap dikenali walau tinggal 1-2 bulan
        if "month_axes" not in [c["name"] for c in g.db.execute("PRAGMA table_info(monitoring_books)")]:
            g.db.execute("ALTER TABLE monitoring_books ADD COLUMN month_axes TEXT NOT NULL DEFAULT '{}'")
    return g.db


@app.teardown_appcontext
def close_db(_exc):
    conn = g.pop("db", None)
    if conn is not None:
        conn.close()


@app.context_processor
def inject_menu():
    # Sidebar: 10 proyek pertama (abjad) + proyek yang sedang dibuka; sisanya lewat Daftar proyek (berhalaman)
    projects = db().execute("SELECT id, name FROM monitoring_projects ORDER BY name").fetchall()
    shown = projects[:MONITORING_PER_PAGE]
    pid = (request.view_args or {}).get("pid")
    shown += [p for p in projects[MONITORING_PER_PAGE:] if p["id"] == pid]
    return {"MENU": MENU, "MONITORING_PROJECTS": shown, "MONITORING_COUNT": len(projects)}


@app.template_filter("rp")
def rp(value):
    return "Rp " + f"{value or 0:,.0f}".replace(",", ".")


@app.template_filter("rp_short")
def rp_short(value):
    """Rupiah ringkas untuk kartu: Rp 4,88 M (miliar), Rp 715 jt, Rp 950 rb."""
    v = value or 0
    for limit, unit in ((1e12, "T"), (1e9, "M"), (1e6, "jt"), (1e3, "rb")):
        if abs(v) >= limit:
            short = f"{v / limit:,.2f}".rstrip("0").rstrip(".").replace(",", "X").replace(".", ",").replace("X", ".")
            return f"Rp {short} {unit}"
    return rp(v)


@app.template_filter("num")
def num(value):
    return f"{value or 0:,.0f}".replace(",", ".")


@app.template_filter("dec")
def dec(value):
    return f"{value or 0:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


@app.template_filter("tanggal")
def tanggal(value):
    """'2026-07-30' -> '30 Jul 2026'."""
    d = parse_date(value)
    return f"{d.day:02d} {MONTHS[d.month - 1]} {d.year}" if d else value


@app.template_filter("waktu")
def waktu(value):
    """'2026-11-21 14:12:36' -> '21 Nov 2026, 14:12'."""
    try:
        d = datetime.strptime(value[:16], "%Y-%m-%d %H:%M")
    except (TypeError, ValueError):
        return value or "-"
    return f"{d.day:02d} {MONTHS[d.month - 1]} {d.year}, {d:%H:%M}"


@app.template_filter("bulan")
def bulan(value):
    """'2026-07' -> 'Jul 2026'."""
    return month_label(value)


@app.template_filter("size")
def human_size(n_bytes):
    units = ["B", "KB", "MB", "GB", "TB", "PB"]
    n = float(n_bytes)
    for unit in units:
        if n < 1024 or unit == units[-1]:
            return f"{dec(n)} {unit}"
        n /= 1024


# ---------- Dashboard ----------

MONTHS = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]


def month_label(ym):
    """'2026-01' -> 'Jan 2026'."""
    year, month = ym.split("-")
    return f"{MONTHS[int(month) - 1]} {year}"


@app.route("/")
def dashboard():
    con = db()
    totals = con.execute(
        "SELECT COUNT(*) AS n, COALESCE(SUM(contract_value),0) AS contract,"
        " COALESCE(SUM(est_cost),0) AS cost FROM projects"
    ).fetchone()
    margin = totals["contract"] - totals["cost"]
    margin_pct = margin / totals["contract"] * 100 if totals["contract"] else 0
    total_sales = con.execute("SELECT COALESCE(SUM(amount),0) FROM sales").fetchone()[0]

    by_status = {r["status"]: r["n"] for r in con.execute(
        "SELECT status, COUNT(*) AS n FROM projects GROUP BY status")}
    monthly = con.execute(
        "SELECT substr(sale_date,1,7) AS month, SUM(amount) AS total"
        " FROM sales GROUP BY month ORDER BY month DESC LIMIT 12").fetchall()[::-1]
    top_products = con.execute(
        "SELECT product, SUM(qty) AS qty, SUM(amount) AS total"
        " FROM sales GROUP BY product ORDER BY total DESC LIMIT 5").fetchall()
    locations = con.execute(
        "SELECT location, SUM(amount) AS total"
        " FROM sales GROUP BY location ORDER BY total DESC LIMIT 6").fetchall()
    recent = con.execute("SELECT * FROM projects ORDER BY id DESC LIMIT 5").fetchall()

    return render_template(
        "dashboard.html", active="dashboard",
        totals=totals, margin=margin, margin_pct=margin_pct, total_sales=total_sales,
        status_labels=STATUSES, status_values=[by_status.get(s, 0) for s in STATUSES],
        month_labels=[month_label(r["month"]) for r in monthly], month_values=[r["total"] for r in monthly],
        top_products=top_products, locations=locations,
        max_location=max([r["total"] for r in locations], default=0),
        recent=recent,
    )


# ---------- Proyek & Kontrak ----------

PDF_FIELDS = {
    "name": "string", "client": "string", "contract_no": "string", "location": "string",
    "contract_value": "number", "start_date": "string", "end_date": "string",
}
PDF_SCHEMA = {
    "type": "object",
    "properties": {k: {"type": [t, "null"]} for k, t in PDF_FIELDS.items()},
    "required": list(PDF_FIELDS),
    "additionalProperties": False,
}
PDF_PROMPT = """Ini dokumen kontrak / SPMK / PO proyek (bisa hasil scan). Ambil data berikut:
- name: nama pekerjaan/proyek
- client: pihak pemberi kerja (yang memberi perintah/kontrak), bukan penerima/subkontraktor
- contract_no: nomor kontrak/SPMK
- location: lokasi proyek (kota/provinsi)
- contract_value: nilai kontrak SEBELUM PPN dalam Rupiah, angka saja (baris "Nilai", bukan "Jumlah" yang sudah termasuk PPN)
- start_date, end_date: tanggal mulai dan berakhir pelaksanaan, format YYYY-MM-DD
Isi null jika tidak tertulis di dokumen. Jangan menebak."""


def read_contract_pdf(pdf_bytes):
    """Kirim PDF ke Gemini API, kembalikan dict field proyek."""
    client = genai.Client()  # simpan di variabel: Client sementara langsung ditutup sebelum request
    interaction = client.interactions.create(
        model="gemini-3.5-flash-lite",
        input=[
            {"type": "document", "mime_type": "application/pdf",
             "data": base64.b64encode(pdf_bytes).decode()},
            {"type": "text", "text": PDF_PROMPT},
        ],
        response_format={"type": "text", "mime_type": "application/json", "schema": PDF_SCHEMA},
    )
    return json.loads(interaction.output_text)


@app.post("/projects/read-pdf")
def project_read_pdf():
    pdf = request.files.get("pdf")
    prefill, error = {}, None
    if not pdf or not pdf.filename.lower().endswith(".pdf"):
        error = "Pilih file PDF."
    else:
        try:
            prefill = read_contract_pdf(pdf.read())
        except json.JSONDecodeError:
            error = "Jawaban Gemini bukan JSON yang valid. Coba lagi."
        except ValueError:  # genai.Client() tidak menemukan API key
            error = "API key belum diatur. Set GEMINI_API_KEY di .env lalu jalankan ulang aplikasi."
        except Exception as e:  # kelas error interactions ada di modul privat SDK
            code = getattr(e, "status_code", None)
            msg = getattr(e, "message", None) or str(e)
            error = f"Gemini API menolak permintaan ({code}): {msg}" if code else f"Gagal membaca PDF: {msg}"
    return projects_page(pre=prefill, pdf_error=error, pdf_read=bool(prefill))


def projects_page(**form_state):
    """Daftar proyek + ringkasan total. Form tambah proyek terbuka bila ada hasil baca PDF atau error."""
    rows = db().execute("SELECT * FROM projects ORDER BY id DESC").fetchall()
    contract = sum(r["contract_value"] for r in rows)
    cost = sum(r["est_cost"] for r in rows)
    summary = {"n": len(rows), "contract": contract, "cost": cost, "margin": contract - cost,
               "margin_pct": (contract - cost) / contract * 100 if contract else 0}
    form_state.setdefault("pre", {})
    open_form = bool(form_state.get("pdf_error") or form_state.get("pdf_read"))
    return render_template("projects.html", active="projects", rows=rows, statuses=STATUSES,
                           rows_json=[dict(r) for r in reversed(rows)],  # chart: terlama ke terbaru
                           summary=summary, open_form=open_form, **form_state)

@app.route("/projects", methods=["GET", "POST"])
def projects():
    con = db()
    if request.method == "POST":
        f = request.form
        con.execute(
            "INSERT INTO projects (name, client, contract_no, location, contract_value,"
            " est_cost, status, start_date, end_date) VALUES (?,?,?,?,?,?,?,?,?)",
            (f["name"], f["client"], f.get("contract_no"), f.get("location"),
             float(f["contract_value"]), float(f["est_cost"]), f["status"], f.get("start_date"),
             f.get("end_date")),
        )
        con.commit()
        return redirect(url_for("projects"))
    return projects_page()


def parse_date(value):
    try:
        return date.fromisoformat(value) if value else None
    except ValueError:
        return None


@app.route("/projects/<int:item_id>")
def project_detail(item_id):
    p = db().execute("SELECT * FROM projects WHERE id=?", (item_id,)).fetchone()
    if p is None:
        abort(404)
    margin = p["contract_value"] - p["est_cost"]
    start, end = parse_date(p["start_date"]), parse_date(p["end_date"])
    timeline = None
    if start and end and end >= start:
        total_days = (end - start).days or 1
        elapsed = min(max((date.today() - start).days, 0), total_days)
        if p["status"] == "Selesai":
            elapsed = total_days
        timeline = {"total": total_days, "elapsed": elapsed, "left": total_days - elapsed,
                    "pct": elapsed / total_days * 100, "not_started": date.today() < start}
    return render_template("project_detail.html", active="projects", p=p, margin=margin, statuses=STATUSES,
                           saved=request.args.get("saved") == "1", edit_error=request.args.get("error", "")[:200],
                           margin_pct=margin / p["contract_value"] * 100 if p["contract_value"] else 0,
                           timeline=timeline)


@app.post("/projects/<int:item_id>/edit")
def project_edit(item_id):
    f = request.form
    try:
        values = (f["name"].strip(), f["client"].strip(), f.get("contract_no"), f.get("location"),
                  float(f["contract_value"]), float(f["est_cost"]), f["status"], f.get("start_date"), f.get("end_date"))
    except (KeyError, ValueError):
        return redirect(url_for("project_detail", item_id=item_id, error="Lengkapi nama, klien, nilai kontrak, dan biaya."))
    if not values[0] or not values[1] or values[4] < 0 or values[5] < 0 or values[6] not in STATUSES:
        return redirect(url_for("project_detail", item_id=item_id, error="Nama, klien, dan status wajib diisi; nilai tidak boleh negatif."))
    con = db()
    updated = con.execute(
        "UPDATE projects SET name=?, client=?, contract_no=?, location=?, contract_value=?, est_cost=?,"
        " status=?, start_date=?, end_date=? WHERE id=?", (*values, item_id)).rowcount
    con.commit()
    if not updated:
        abort(404)
    return redirect(url_for("project_detail", item_id=item_id, saved=1))


@app.post("/projects/<int:item_id>/start")
def project_start(item_id):
    """Checklist mulai: centang -> Berjalan, hapus centang -> Belum Mulai. Hanya untuk dua status itu."""
    status = "Berjalan" if request.form.get("started") == "1" else "Belum Mulai"
    con = db()
    con.execute("UPDATE projects SET status=? WHERE id=? AND status IN ('Belum Mulai', 'Berjalan')", (status, item_id))
    con.commit()
    return redirect(url_for("projects"))


@app.post("/projects/<int:item_id>/delete")
def project_delete(item_id):
    db().execute("DELETE FROM projects WHERE id=?", (item_id,))
    db().commit()
    return redirect(url_for("projects"))


# ---------- Pencarian (command palette) ----------

@app.get("/search")
def search():
    """Cari proyek, riwayat hitungan storage, dan penjualan. q kosong = item terbaru (tampilan awal palette)."""
    q = request.args.get("q", "").strip()[:80]
    like = "%" + q.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_") + "%"
    where = lambda cols: (" WHERE " + " OR ".join(f"{c} LIKE ? ESCAPE '\\'" for c in cols)) if q else ""
    args = lambda cols: [like] * len(cols) if q else []
    con = db()
    cols = ["name", "client", "contract_no", "location"]
    projects = [{"label": r["name"], "desc": f"{r['client']} · {r['status'].capitalize()}",
                 "url": url_for("project_detail", item_id=r["id"])}
                for r in con.execute(f"SELECT * FROM projects{where(cols)} ORDER BY id DESC LIMIT 5", args(cols))]
    cols = ["name", "summary"]
    logs = [{"label": r["name"] or r["summary"], "desc": f"{human_size(r['total'])} · {r['used_at'][:10]}",
             "url": url_for("storage") + "?" + r["query"]}
            for r in con.execute(f"SELECT * FROM storage_logs{where(cols)} ORDER BY used_at DESC LIMIT 5", args(cols))]
    cols = ["product", "location"]
    sales = [{"label": r["product"], "desc": f"{r['location']} · {r['sale_date']} · {rp(r['amount'])}",
              "url": url_for("sales")}
             for r in con.execute(f"SELECT * FROM sales{where(cols)} ORDER BY sale_date DESC, id DESC LIMIT 5", args(cols))] if q else []
    return jsonify({"projects": projects, "logs": logs, "sales": sales})


# ---------- Sales ----------

@app.route("/sales", methods=["GET", "POST"])
def sales():
    con = db()
    if request.method == "POST":
        f = request.form
        con.execute(
            "INSERT INTO sales (sale_date, product, location, qty, amount) VALUES (?,?,?,?,?)",
            (f["sale_date"], f["product"], f["location"], int(f["qty"]), float(f["amount"])),
        )
        con.commit()
        return redirect(url_for("sales"))
    rows = con.execute("SELECT * FROM sales ORDER BY sale_date DESC, id DESC").fetchall()
    return render_template("sales.html", active="sales", rows=rows)


@app.post("/sales/<int:item_id>/delete")
def sale_delete(item_id):
    db().execute("DELETE FROM sales WHERE id=?", (item_id,))
    db().commit()
    return redirect(url_for("sales"))


# ---------- Kalkulator Storage ----------

# (batas record/detik, vCPU, RAM GB, catatan) — aturan praktis, bukan hasil benchmark.
SERVER_TIERS = [
    (1_000, 4, 8, "Satu server database sudah cukup."),
    (10_000, 8, 16, "Gunakan SSD dan batch insert."),
    (100_000, 16, 32, "Gunakan NVMe dan database time-series (mis. TimescaleDB)."),
    (math.inf, 32, 64, "Pertimbangkan cluster / beberapa node dan kompresi data."),
]


# Satu jenis sensor = 1 tabel raw + 1 tabel statistik (pola marunda_dashboard/init_db.sql).
# channels: kolom ukur (double precision) di tabel raw; tabel statistik menyimpan min/max/avg tiap channel.
# Urutan = urutan warna chart; warna mengikuti jenis, bukan peringkat.
# "usulan": jenis yang belum punya tabel di skema Marunda, kolomnya asumsi.
# color: token --cat-N di static/css/arc.css (versi terang & gelap, keduanya lolos validasi palet).
SENSOR_TYPES = {
    "strain_gauge": {"label": "Strain gauge", "color": "var(--cat-1)", "channels": ["strain_ue", "temp_c"],
                     "raw": "strain", "stat": "strain_statistic"},
    "tiltmeter": {"label": "Tiltmeter", "color": "var(--cat-2)", "channels": ["angle_x", "angle_y"],
                  "raw": "tiltmeter", "stat": "tiltmeter_statistic"},
    "accelerometer": {"label": "Accelerometer", "color": "var(--cat-3)", "channels": ["ax", "ay", "az"],
                      "raw": "accelerometers", "stat": "accelerometers_statistic (usulan)"},
    "weather_station": {"label": "Weather station", "color": "var(--cat-4)",
                        "channels": ["wind_speed", "wind_direction", "air_temperature", "humidity"],
                        "raw": "weather_station", "stat": "weather_station_statistic"},
    "anemometer_2d": {"label": "Anemometer 2D", "color": "var(--cat-5)", "channels": ["wind_speed", "wind_direction"],
                      "raw": "anm2d", "stat": "anm2d_statistic"},
    "anemometer_3d": {"label": "Anemometer 3D", "color": "var(--cat-6)", "channels": ["wind_u", "wind_v", "wind_w"],
                      "raw": "anm3d (usulan)", "stat": "anm3d_statistic (usulan)"},
    "jointmeter": {"label": "Jointmeter", "color": "var(--cat-7)", "channels": ["displacement_mm", "temp_c"],
                   "raw": "jointmeter (usulan)", "stat": "jointmeter_statistic (usulan)"},
}

# Byte per baris = data + index, diukur di PostgreSQL dengan tabel berskema Marunda (pg_relation_size +
# pg_indexes_size): raw strain (2 channel) 109 + 119 byte, strain_statistic (6 kolom) 126 + 62 byte.
# Tiap kolom double menambah 8 byte. Belum termasuk kompresi TimescaleDB.
RAW_ROW_BASE, STAT_ROW_BASE, BYTES_PER_COLUMN = 212, 140, 8
STATS_PER_CHANNEL = 3  # min, max, avg

INTERVAL_UNITS = {"s": 1, "min": 60}  # satuan sampling interval -> detik


def raw_days_equivalent(days, compress_after, compress_pct):
    """Hari-setara tabel raw setelah `days` hari: data > compress_after hari dikompres (hypertable Marunda)."""
    return min(days, compress_after) + max(days - compress_after, 0) * (1 - compress_pct / 100)


def calc_storage(rows, retention_days, stat_minutes, compress_after=30, compress_pct=0):
    """rows: list of (type, qty, interval_s, channels). Hasil per jenis (raw + statistik) + total.
    Kompresi TimescaleDB hanya untuk tabel raw; tabel statistik bukan hypertable."""
    by_type = {}
    for kind, qty, interval_s, channels in rows:
        raw_rate = qty / interval_s  # baris raw per detik
        stat_rate = qty / (stat_minutes * 60)  # baris statistik per detik
        raw_row = RAW_ROW_BASE + BYTES_PER_COLUMN * channels
        stat_row = STAT_ROW_BASE + BYTES_PER_COLUMN * STATS_PER_CHANNEL * channels
        meta = SENSOR_TYPES[kind]
        t = by_type.setdefault(kind, {"label": meta["label"], "color": meta["color"], "raw_table": meta["raw"],
                                      "stat_table": meta["stat"], "qty": 0, "channels": set(), "rate": 0.0,
                                      "raw_day": 0.0, "stat_day": 0.0})
        t["qty"] += qty
        t["channels"].add(channels)
        t["rate"] += raw_rate + stat_rate
        t["raw_day"] += raw_rate * 86_400 * raw_row
        t["stat_day"] += stat_rate * 86_400 * stat_row
    types = [by_type[k] for k in SENSOR_TYPES if k in by_type]  # urutan tetap
    raw_days = raw_days_equivalent(retention_days, compress_after, compress_pct)
    for t in types:
        t["channels"] = ", ".join(str(c) for c in sorted(t["channels"]))
        t["per_day"] = t["raw_day"] + t["stat_day"]
        t["raw_uncompressed"] = t["raw_day"] * retention_days
        t["raw_total"] = t["raw_day"] * raw_days
        t["stat_total"] = t["stat_day"] * retention_days
        t["total"] = t["raw_total"] + t["stat_total"]
    rate = sum(t["rate"] for t in types)
    per_day = sum(t["per_day"] for t in types)
    raw_day = sum(t["raw_day"] for t in types)
    stat_day = sum(t["stat_day"] for t in types)

    def storage_at(days):
        return raw_day * raw_days_equivalent(days, compress_after, compress_pct) + stat_day * days

    total = storage_at(retention_days)
    _, cpu, ram, note = next(t for t in SERVER_TIERS if rate < t[0])
    return {
        "types": types,
        "rate": rate,
        "raw_day": raw_day,
        "stat_day": stat_day,
        "raw_uncompressed": raw_day * retention_days,
        "raw_total": sum(t["raw_total"] for t in types),
        "stat_total": sum(t["stat_total"] for t in types),
        "saved": raw_day * retention_days - sum(t["raw_total"] for t in types),
        "per_day": per_day,
        "per_month": storage_at(30),
        "per_year": storage_at(365),
        "total": total,
        "disk": total * 1.3,  # cadangan ruang kosong 30%
        "cpu": cpu,
        "ram": ram,
        "note": note,
    }


@app.route("/storage")
def storage():
    a = request.args
    rows = list(zip(a.getlist("type"), a.getlist("qty"), a.getlist("interval"), a.getlist("unit"),
                    a.getlist("channels")))
    # Kunci riwayat = semua input kecuali nama, jadi mengganti nama tidak membuat baris baru
    query = urlencode([(k, v) for k, v in a.items(multi=True) if k != "name"])
    inputs = {
        "name": a.get("name", "").strip()[:80],
        "rows": rows or [("strain_gauge", 10, 0.01, "s", len(SENSOR_TYPES["strain_gauge"]["channels"]))],
        "retention_days": a.get("retention_days", 365, type=int),
        "stat_minutes": a.get("stat_minutes", 10, type=int),
        "compress_after": a.get("compress_after", 30, type=int),
        "compress_pct": a.get("compress_pct", 90.0, type=float),
    }
    result = error = None
    if query:
        try:
            parsed = [(k, int(q), float(i) * INTERVAL_UNITS[u], int(c)) for k, q, i, u, c in rows]
        except (ValueError, KeyError):
            parsed = None
        if not parsed:
            error = "Lengkapi semua input dengan angka."
        elif any(k not in SENSOR_TYPES or q <= 0 or i <= 0 or c <= 0 for k, q, i, c in parsed) \
                or inputs["retention_days"] <= 0 or inputs["stat_minutes"] <= 0:
            error = "Jumlah, interval, channel, retensi, dan periode statistik harus lebih besar dari 0."
        elif inputs["compress_after"] < 0 or not 0 <= inputs["compress_pct"] < 100:
            error = "Kompres setelah minimal 0 hari, hemat kompresi 0 sampai di bawah 100%."
        else:
            result = calc_storage(parsed, inputs["retention_days"], inputs["stat_minutes"],
                                  inputs["compress_after"], inputs["compress_pct"])
            log_storage(query, inputs["name"], result, inputs["retention_days"])
    con = db()
    if query and not inputs["name"]:  # dimuat dari riwayat tanpa nama di URL: ambil nama tersimpan
        saved = con.execute("SELECT name FROM storage_logs WHERE query=?", (query,)).fetchone()
        inputs["name"] = (saved and saved["name"]) or ""
    logs = con.execute("SELECT * FROM storage_logs ORDER BY used_at DESC, id DESC LIMIT 20").fetchall()
    return render_template("storage.html", active="storage", inputs=inputs, result=result, error=error,
                           sensor_types=SENSOR_TYPES, raw_row_base=RAW_ROW_BASE, stat_row_base=STAT_ROW_BASE,
                           logs=logs, current_query=query)


def log_storage(query, name, result, retention_days):
    """Simpan hitungan ke riwayat. Input yang sama persis tidak diduplikasi: waktunya diperbarui (terakhir dipakai);
    nama baru menggantikan nama lama, nama kosong mempertahankan nama lama."""
    summary = ", ".join(f"{t['qty']} {t['label'].lower()}" for t in result["types"]) + f" · {retention_days} hari"
    con = db()
    con.execute(
        "INSERT INTO storage_logs (used_at, query, summary, total, name) VALUES (?,?,?,?,?)"
        " ON CONFLICT(query) DO UPDATE SET used_at = excluded.used_at, total = excluded.total,"
        " name = COALESCE(NULLIF(excluded.name, ''), storage_logs.name)",
        (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), query, summary, result["total"], name),
    )
    con.commit()


@app.post("/storage/logs/<int:log_id>/delete")
def storage_log_delete(log_id):
    db().execute("DELETE FROM storage_logs WHERE id=?", (log_id,))
    db().commit()
    back = request.form.get("back", "")
    return redirect(url_for("storage") + (f"?{back}" if back else ""))


# ---------- Monitoring Project ----------
# Dua jenis proyek: 'manual' (isian dashboard + cost bulanan) dan 'xlsx' (snapshot file Excel di database;
# sel input bisa diedit, rumus dihitung ulang dengan pycel).


def excel_number(v, decimals, thousands=True):
    """Format angka gaya Indonesia: 1.234.567,89."""
    s = f"{v:,.{decimals}f}" if thousands else f"{v:.{decimals}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def excel_value(v, fmt):
    """Tampilkan nilai sel kira-kira seperti di Excel, mengikuti number format sel."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (datetime, date)):
        f = re.sub(r"\[\$[^\]]*\]", "", fmt).lower()
        month = MONTHS[v.month - 1]
        if "y" not in f:
            return month
        return f"{v.day:02d} {month} {v.year}" if "d" in f else f"{month} {v.year}"
    if isinstance(v, (int, float)):
        if fmt in ("General", "@"):
            return f"{v:.10g}".replace(".", ",")
        section = fmt.split(";")[0]
        if "%" in fmt:
            m = re.search(r"0\.(0+)", section)
            return excel_number(v * 100, len(m.group(1)) if m else 0) + "%"
        if v == 0 and '"-"' in fmt:
            return "-"
        m = re.search(r"0\.(0+)", section)
        s = excel_number(abs(v), len(m.group(1)) if m else 0, thousands="#,##" in section)
        if "Rp" in fmt:
            s = "Rp " + s
        if v < 0:
            s = f"({s})" if "\\(" in fmt else "-" + s
        return s
    return str(v)


def read_sheet(ws, value=lambda c: c.value, max_col=None, max_row=None, edited=(), pum=None):
    """Sheet -> (baris {n: nomor baris Excel, cells}, lebar kolom px) untuk tabel HTML, dibatasi sampai max_col/max_row. Merged cell jadi
    rowspan/colspan, baris tersembunyi dilewati. value(cell) memberi nilai yang ditampilkan (hasil hitung ulang rumus).
    Gaya ikut Excel: border per sisi, warna isi, teks putih di atas isi gelap, font besar, lebar kolom.
    pum = {sel: Rp} angka PUM yang sudah ditambahkan ke sel; isian edit menampilkan nilai tanpa PUM."""
    cells = [c for row in ws.iter_rows(max_row=max_row or ws.max_row, max_col=max_col or ws.max_column)
             for c in row if c.value is not None and str(c.value).strip() != ""]
    if not cells:
        return [], []
    r0, r1 = min(c.row for c in cells), max(c.row for c in cells)
    c0, c1 = min(c.column for c in cells), max(c.column for c in cells)
    hidden = {r for r, d in ws.row_dimensions.items() if d.hidden}
    spans, covered = {}, set()
    for rng in ws.merged_cells.ranges:
        top, left = max(rng.min_row, r0), max(rng.min_col, c0)
        bottom, right = min(rng.max_row, r1), min(rng.max_col, c1)
        if top > bottom or left > right:
            continue
        rows = sum(1 for r in range(top, bottom + 1) if r not in hidden)
        spans[(top, left)] = (rows, right - left + 1, bottom, right)
        covered |= {(r, c) for r in range(top, bottom + 1) for c in range(left, right + 1)} - {(top, left)}
    out, date_cols = [], {}  # date_cols: baris -> kolom berisi tanggal (judul bulan)
    for r in range(r0, r1 + 1):
        if r in hidden:
            continue
        row = []
        for c in range(c0, c1 + 1):
            if (r, c) in covered:
                continue
            cell = ws.cell(r, c)
            v = value(cell)
            formula = cell.value if is_formula(cell) else None
            rowspan, colspan, bottom, right = spans.get((r, c), (1, 1, r, c))
            is_num = isinstance(v, (int, float)) and not isinstance(v, bool)
            align = cell.alignment.horizontal
            if is_num:
                align = "right"  # angka selalu rata kanan agar digit sejajar
            elif align not in ("left", "center", "right"):
                align = "right" if isinstance(v, (datetime, date)) else None
            text = excel_value(v, cell.number_format)
            if isinstance(v, (datetime, date)) and colspan == 1:
                date_cols.setdefault(r, set()).add(c)
            # teks yang di Excel meluber ke sel kosong di kanannya (judul, label bagian): tidak menentukan lebar kolom
            right_empty = c == c1 or ((r, c + 1) not in covered and ws.cell(r, c + 1).value in (None, ""))
            # sel bergaris kanan ada di dalam tabel: teksnya tidak meluber, jadi ikut menentukan lebar kolom
            boxed = (cell.border.right is not None and cell.border.right.style in BORDER_WEIGHT) or \
                (c < c1 and ws.cell(r, c + 1).border.left is not None and ws.cell(r, c + 1).border.left.style in BORDER_WEIGHT)
            spill = bool(text) and not is_num and colspan == 1 and right_empty and not boxed
            # merged cell: border kanan/bawah ada di sel tepi kanan/bawah range-nya
            sides = (cell.border.left, ws.cell(r, right).border.right, cell.border.top, ws.cell(bottom, c).border.bottom)
            fill = cell.fill.fgColor if cell.fill.fill_type == "solid" else None
            font = cell.font.color
            row.append({"text": text, "rowspan": rowspan, "colspan": colspan, "col": c, "num": is_num, "spill": spill,
                        "bold": bool(cell.font.b), "align": align,
                        "ref": cell.coordinate, "formula": formula, "pum": (pum or {}).get(cell.coordinate),
                        "input": input_text(v - (pum or {}).get(cell.coordinate, 0) if is_num else v),
                        "edited": cell.coordinate in edited,
                        "borders": "".join(f" xb-{k}{BORDER_WEIGHT[side.style]}" for k, side in zip("lrtb", sides)
                                           if side is not None and side.style in BORDER_WEIGHT),
                        "fill": None if fill is not None and fill.type == "rgb" and fill.rgb[-6:].upper() in HIDDEN_FILLS
                        else ("#" + fill.rgb[-6:]) if fill is not None and fill.type == "rgb" else ("theme" if fill is not None else None),
                        "on_fill": font is not None and ((font.type == "rgb" and font.rgb[-6:] == "FFFFFF") or (font.type == "theme" and font.theme == 0)),
                        "big": (cell.font.sz or 11) >= 13})
        out.append({"n": r, "cells": row})
    # Lebar kolom sebanding isi: karakter terpanjang di kolom (tanpa teks meluber & merged cell), dibatasi min/maks
    longest = {c: 0 for c in range(c0, c1 + 1)}
    for row in out:
        for cell in row["cells"]:
            if cell["colspan"] == 1 and not cell["spill"]:
                longest[cell["col"]] = max(longest[cell["col"]], len(cell["text"]))
    width = {c: min(COL_MAX_PX, max(COL_MIN_PX, round(n * 7.5) + 20)) if n else COL_EMPTY_PX for c, n in longest.items()}
    # merged cell: kekurangan lebar dibagi rata ke kolom yang dicakupnya.
    # teks meluber: muat di sel kosong di kanannya; bila kurang (mis. kolom terakhir), kolom sel itu sendiri dilebarkan.
    for row in out:
        filled = {c["col"] for c in row["cells"] if c["text"] or c["colspan"] > 1}
        for cell in row["cells"]:
            if not cell["text"] or not (cell["colspan"] > 1 or cell["spill"]):
                continue
            need = min(COL_MAX_PX, round(len(cell["text"]) * 7.5) + 20)
            if cell["colspan"] > 1:
                cols = range(cell["col"], cell["col"] + cell["colspan"])
                short = need - sum(width[c] for c in cols)
                for c in cols if short > 0 else ():
                    width[c] += math.ceil(short / cell["colspan"])
            else:
                end = cell["col"] + 1
                while end <= c1 and end not in filled:
                    end += 1
                short = need - sum(width[c] for c in range(cell["col"], end))
                if short > 0:
                    width[cell["col"]] += short
    # Deretan judul bulan (>= 3 tanggal dalam satu baris): semua kolomnya selebar kolom bulan terlebar
    for cols in date_cols.values():
        if len(cols) >= 3:
            widest = max(width[c] for c in cols)
            for c in cols:
                width[c] = widest
    # Kolom di bawah satu judul merge yang masing-masing punya sub-judul (mis. Ammount -> Teguh | Panji | ...):
    # disamakan lebarnya supaya proporsional
    texts = {(row["n"], cell["col"]): cell["text"] for row in out for cell in row["cells"]}
    for (top, left), (_rows, colspan, bottom, right) in spans.items():
        if colspan > 1 and all(texts.get((bottom + 1, c)) for c in range(left, right + 1)):
            widest = max(width[c] for c in range(left, right + 1))
            for c in range(left, right + 1):
                width[c] = widest
    return out, list(width.values())


COL_MIN_PX, COL_MAX_PX, COL_EMPTY_PX = 44, 380, 16  # batas lebar kolom tabel sheet (teks lebih panjang dibungkus)


HIDDEN_FILLS = {"FFFF00"}  # warna isi Excel yang tidak ditampilkan (kuning penanda di file)

# Ketebalan border Excel -> kelas CSS xb-<sisi><1|2>
BORDER_WEIGHT = {"hair": 1, "thin": 1, "dotted": 1, "dashed": 1, "dashDot": 1, "dashDotDot": 1,
                 "medium": 2, "mediumDashed": 2, "mediumDashDot": 2, "mediumDashDotDot": 2, "thick": 2, "double": 2,
                 "slantDashDot": 2}


def is_formula(cell):
    return isinstance(cell.value, str) and cell.value.startswith("=")


def input_text(v):
    """Nilai sel untuk isian form edit: angka tanpa pemisah ribuan, tanggal YYYY-MM-DD."""
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def ref_cells(wb, ref):
    """"'Sheet'!$E$10:$M$10" -> daftar sel (urut baris lalu kolom)."""
    sheet, rng = ref.rsplit("!", 1)
    cells = wb[sheet.strip("'").replace("''", "'")][rng.replace("$", "")]
    if not isinstance(cells, tuple):
        return [cells]
    return [c for row in cells for c in (row if isinstance(row, tuple) else (row,))]


def data_ref(source):
    """Rujukan sel dari sumber data seri chart (kategori/nilai/x/y), None bila bukan rujukan."""
    for kind in ("numRef", "strRef"):
        r = getattr(source, kind, None) if source is not None else None
        if r is not None:
            return r.f
    return None


def read_charts(ws, wb, value=lambda c: c.value, inside=lambda c: True):
    """Chart Excel di sheet -> {title, labels, series:[{label, values}]} dari sel yang dirujuk tiap seri.
    Seri dari bar chart ditandai bar=True (batang); line dan scatter digambar sebagai garis.
    Chart yang merujuk sel di luar batas (inside(cell) False) dilewati."""
    charts = []
    for ch in ws._charts:
        try:
            labels, series = None, []
            for sub in getattr(ch, "_charts", None) or [ch]:  # chart kombinasi (bar + line) punya beberapa sub-chart
                for s in sub.series:
                    y = data_ref(s.val) or data_ref(s.yVal)
                    if not y:
                        continue
                    refs = [ref_cells(wb, y)] + [ref_cells(wb, r) for r in (data_ref(s.cat) or data_ref(s.xVal),) if r]
                    if not all(inside(c) for cells in refs for c in cells):
                        raise ValueError("chart di luar batas")
                    if labels is None and len(refs) > 1:
                        labels = [excel_value(value(c), c.number_format) for c in refs[1]]
                    name = None
                    if s.tx is not None:
                        name = value(ref_cells(wb, s.tx.strRef.f)[0]) if s.tx.strRef is not None else s.tx.v
                    values = [v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0
                              for v in map(value, refs[0])]
                    series.append({"label": str(name or f"Seri {len(series) + 1}"), "values": values,
                                   "bar": type(sub).__name__.startswith("BarChart")})
            if not series:
                continue
            n = max(len(s["values"]) for s in series)
            try:
                title = "".join(r.t for p in ch.title.tx.rich.p for r in (p.r or []))
            except AttributeError:
                title = ""
            charts.append({"title": title or "Chart", "labels": labels or [str(i + 1) for i in range(n)],
                           "series": series})
        except (KeyError, ValueError, IndexError):  # rujukan ke sheet/rentang yang tidak ada atau di luar batas
            continue
    return charts


# ---------- Monitoring dari file Excel: snapshot di database + hitung ulang rumus (pycel) ----------

# Sheet yang tampil dan bisa diedit: (kata kunci judul sheet, kolom terakhir, baris terakhir di file asli
# (None = sampai akhir), sheet daftar = baris bisa ditambah/dihapus). Sheet lain bukan daftar: barisnya kode biaya
# yang dicocokkan antar-sheet (INDEX/MATCH), jadi tidak boleh disisipi.
MONITORING_LAYOUT = [
    ("breakdown cost", "I", 262, True), ("initial budget", "E", 49, False), ("cost planning", "P", 49, False),
    ("actual cost", "N", 54, False), ("rekap progress", "P", 49, False), ("material support", "G", None, True),
    ("tools", "G", None, True),
]
# Tidak tampil sebagai tab, tetap disimpan karena dirujuk rumus. Data Project diedit lewat form Edit data proyek.
MONITORING_INPUT_SHEETS = ["data project", "contract"]
_books = {}  # project_id -> {key, wb, xl, applied}: workbook + compiler pycel yang sudah dimuat


def find_sheet(wb, key):
    return next((ws for ws in wb.worksheets
                 if key in ws.title.lower() and not ws.title.lower().startswith("arsip")), None)


def make_snapshot(source):
    """File Excel (path atau stream) -> bytes xlsx berisi hanya sheet monitoring (+ sheet input). Sheet lain dibuang."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(source)  # rumus, bukan nilai cache
    keep = []
    for key in [layout[0] for layout in MONITORING_LAYOUT] + MONITORING_INPUT_SHEETS:
        ws = find_sheet(wb, key)
        if ws is None:
            raise ValueError(f"sheet '{key}' tidak ditemukan")
        keep.append(ws.title)
    for ws in list(wb.worksheets):
        if ws.title not in keep:
            wb.remove(ws)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def save_snapshot(pid, source):
    con = db()
    con.execute("INSERT OR REPLACE INTO monitoring_books (project_id, xlsx, imported_at) VALUES (?,?,?)",
                (pid, make_snapshot(source), datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    con.execute("DELETE FROM monitoring_edits WHERE project_id=?", (pid,))
    con.commit()


def load_book(pid):
    """Snapshot proyek + editan sel -> {wb, xl, imported_at, edits}. None bila belum diimpor."""
    row = db().execute("SELECT xlsx, imported_at, extra_rows, extra_cols, month_axes FROM monitoring_books WHERE project_id=?",
                       (pid,)).fetchone()
    if row is None:
        return None
    edits = {(r["sheet"], r["cell"]): json.loads(r["value"]) for r in
             db().execute("SELECT sheet, cell, value FROM monitoring_edits WHERE project_id=?", (pid,))}
    book = _books.get(pid)
    if not book or book["key"] != row["imported_at"]:
        book = _books[pid] = {"key": row["imported_at"], "wb": openpyxl.load_workbook(io.BytesIO(row["xlsx"])),
                              "xl": None, "applied": None, "pum_extra": {}, "pum_unmatched": []}
    pums = pum_actuals(pid)
    # Editan / PUM berubah -> ditulis ke salinan snapshot, lalu compiler dibuat dari salinan itu: semua rumus
    # dihitung dari input terbaru. (set_value pycel tidak membatalkan rumus yang merujuk rentang, mis. INDEX/MATCH.)
    if book["applied"] != (edits, pums):
        wb = openpyxl.load_workbook(io.BytesIO(row["xlsx"]))
        for (sheet, cell), value in edits.items():
            wb[sheet][cell].value = value
            if isinstance(value, str):
                wb[sheet][cell].data_type = "s"  # teks berawalan "=" tetap teks, bukan rumus
        extra, unmatched = add_pum_to_actual(wb, compile_workbook(wb), pums) if pums else ({}, [])
        book.update(xl=compile_workbook(wb), applied=(edits, pums), pum_extra=extra, pum_unmatched=unmatched)
    return {**book, "imported_at": row["imported_at"], "edits": edits, "extra_rows": json.loads(row["extra_rows"]),
            "extra_cols": json.loads(row["extra_cols"]), "axes_meta": json.loads(row["month_axes"]), "pid": pid}


def compile_workbook(wb):
    """Workbook openpyxl -> ExcelCompiler pycel (pycel membaca dari file)."""
    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    try:
        os.close(fd)
        wb.save(tmp)
        return ExcelCompiler(filename=tmp, plugins=["excel_functions"])  # DATEDIF belum ada di pycel
    finally:
        os.unlink(tmp)


def add_pum_to_actual(wb, xl, pums):
    """Tambahkan total PUM {(kategori, 'YYYY-MM'): Rp} ke sel input sheet Actual Cost (baris kategori x kolom bulan).
    xl = compiler sebelum PUM, untuk membaca tanggal di baris judul bulan. -> ({(sheet, sel): Rp}, [tak tertampung])"""
    ws = find_sheet(wb, "actual cost")
    if ws is None:
        return {}, [(cat, ym, amt) for (cat, ym), amt in pums.items()]
    rows = {str(ws.cell(r, 2).value).strip().lower(): r for r in range(1, ws.max_row + 1) if ws.cell(r, 2).value}
    known = [r for cat, r in rows.items() if any(cat == c.lower() for c, _m in pums)]
    months = {}
    for r in range(min(known or [1]) - 1, 0, -1):  # baris judul bulan terdekat di atas baris kategori
        for c in range(3, ws.max_column + 1):
            v = book_value({"xl": xl}, ws.cell(r, c))
            if isinstance(v, datetime):
                months.setdefault(v.strftime("%Y-%m"), c)
        if months:
            break
    extra, unmatched = {}, []
    for (cat, ym), amt in pums.items():
        r, c = rows.get(cat.lower()), months.get(ym)
        cell = ws.cell(r, c) if r and c else None
        if cell is None or is_formula(cell) or isinstance(cell, MergedCell):
            unmatched.append((cat, ym, amt))
            continue
        base = cell.value if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool) else 0
        cell.value = base + amt
        extra[(ws.title, cell.coordinate)] = amt
    return extra, unmatched


def book_value(book, cell):
    """Nilai sel setelah editan dan hitung ulang rumus. Angka di sel berformat tanggal -> datetime."""
    try:
        v = book["xl"].evaluate(f"'{cell.parent.title}'!{cell.coordinate}")
    except Exception:  # rumus yang tidak bisa dihitung
        return "#ERR"
    if isinstance(v, (int, float)) and not isinstance(v, bool) and is_date_format(cell.number_format):
        return from_excel(v)
    return None if v == "" else v


def book_layout(book):
    """[(worksheet, kolom terakhir (angka), baris terakhir, sheet daftar)] untuk sheet yang tampil, urut
    MONITORING_LAYOUT. Baris terakhir ikut bertambah/berkurang sesuai baris yang disisipkan/dihapus."""
    out = []
    for key, col, row, is_list in MONITORING_LAYOUT:
        ws = find_sheet(book["wb"], key)
        if ws is not None:
            last = row + book["extra_rows"].get(ws.title, 0) if row else ws.max_row
            out.append((ws, column_index_from_string(col) + book["extra_cols"].get(ws.title, 0), last, is_list))
    return out


# ---- Tambah / hapus baris di sheet daftar: rujukan rumus di semua sheet ikut digeser seperti di Excel ----

CELL_REF = re.compile(r"^(?P<sheet>(?:'(?:[^']|'')+'|[^'!:]+)!)?(?P<c1>\$?[A-Z]{1,3})(?P<a1>\$?)(?P<r1>\d+)"
                      r"(?::(?P<c2>\$?[A-Z]{1,3})(?P<a2>\$?)(?P<r2>\d+))?$")


def shift_ref(token, own_sheet, target, at, delta, axis="row", track_last=False, band=None):
    """Satu rujukan sel/rentang setelah baris (axis="row") atau kolom (axis="col") ke-`at` di sheet `target`
    disisipkan (delta=1, yang baru = at) atau dihapus (delta=-1). Rujukan ke yang dihapus -> #REF!. Rentang di
    sheet yang sama yang berakhir tepat sebelum sisipan ikut diperlebar, supaya SUM daftar / bulan mencakupnya."""
    # band (kolom): hanya rujukan yang barisnya di dalam pita (baris_atas, baris_bawah) yang bergeser.
    # track_last (kolom bulan): rujukan kolom absolut ($M) ke kolom terakhir adalah jangkar "bulan terakhir", mis.
    # MIN($O11,$M$10) di distribusi rencana; jangkar ikut pindah ke kolom terakhir yang baru (sisip/hapus di ujung).
    m = CELL_REF.match(token)
    if not m:
        return token  # nama, seluruh kolom ($B:$B), rujukan eksternal: tidak bergeser
    sheet = m["sheet"][:-1].strip("'").replace("''", "'") if m["sheet"] else own_sheet
    if sheet != target:
        return token
    if band and axis == "col" and not (band[0] <= int(m["r1"]) <= band[1] and band[0] <= int(m["r2"] or m["r1"]) <= band[1]):
        return token
    if axis == "row":
        i1, i2 = int(m["r1"]), int(m["r2"] or m["r1"])
    else:
        i1 = column_index_from_string(m["c1"].lstrip("$"))
        i2 = column_index_from_string((m["c2"] or m["c1"]).lstrip("$"))
    anchor = track_last and axis == "col" and not m["r2"] and m["c1"].startswith("$")
    if delta > 0:
        n1, n2 = i1 + (i1 >= at), i2 + (i2 >= at)
        if m["r2"] and i1 < i2 == at - 1 and (own_sheet == target or track_last):
            n2 = i2 + 1
        if anchor and i1 == at - 1:
            n1 = n2 = at
    elif i1 == i2 == at:
        if not anchor:
            return "#REF!"
        n1 = n2 = at - 1
    else:
        n1, n2 = i1 - (i1 > at), i2 - (i2 >= at)
    if axis == "row":
        out = f"{m['sheet'] or ''}{m['c1']}{m['a1']}{n1}"
        return out + (f":{m['c2']}{m['a2']}{n2}" if m["r2"] else "")
    col = lambda orig, n: ("$" if orig.startswith("$") else "") + get_column_letter(n)
    out = f"{m['sheet'] or ''}{col(m['c1'], n1)}{m['a1']}{m['r1']}"
    return out + (f":{col(m['c2'], n2)}{m['a2']}{m['r2']}" if m["r2"] else "")


def shift_formula(formula, own_sheet, target, at, delta, axis="row", track_last=False, band=None):
    tok = Tokenizer(formula)
    for t in tok.items:
        if t.type == "OPERAND" and t.subtype == "RANGE":
            t.value = shift_ref(t.value, own_sheet, target, at, delta, axis, track_last, band)
    return "=" + "".join(t.value for t in tok.items)


def formula_cells(wb):
    """(worksheet, sel, teks rumus) untuk semua rumus biasa dan array formula."""
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if is_formula(c):
                    yield ws, c, c.value
                elif isinstance(c.value, ArrayFormula):
                    yield ws, c, c.value.text


def set_formula(cell, text):
    if isinstance(cell.value, ArrayFormula):
        cell.value.text = text
    else:
        cell.value = text


def row_references(wb, ws, row):
    """Sel di luar baris `row` yang rumusnya merujuk baris itu (akan jadi #REF! bila baris dihapus)."""
    found = []
    for w, c, text in formula_cells(wb):
        if w is ws and c.row == row:
            continue
        if "#REF!" in shift_formula(text, w.title, ws.title, row, -1) and "#REF!" not in text:
            found.append(f"{w.title}!{c.coordinate}")
    return found


def change_rows(wb, ws, at, delta):
    """Sisipkan baris baru `at` (delta=1, rumus & format disalin dari baris di atasnya, sel input kosong)
    atau hapus baris `at` (delta=-1). Rumus semua sheet, merged cell, dan tinggi/sembunyi baris ikut digeser.
    (Validasi data, conditional formatting, dan chart tidak digeser: sheet daftar tidak memilikinya.)"""
    for w, c, text in formula_cells(wb):
        set_formula(c, shift_formula(text, w.title, ws.title, at, delta))
    moved = lambda r: r + (r >= at) if delta > 0 else (None if r == at else r - (r > at))
    merged = [rng.coord for rng in ws.merged_cells.ranges]
    for coord in merged:
        ws.unmerge_cells(coord)
    dims = list(ws.row_dimensions.items())
    if delta > 0:
        ws.insert_rows(at)
        for col in range(1, ws.max_column + 1):
            src, dst = ws.cell(at - 1, col), ws.cell(at, col)
            dst._style = copy.copy(src._style)
            if is_formula(src):
                dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
    else:
        ws.delete_rows(at)
    ws.row_dimensions.clear()
    for r, d in dims:
        if moved(r):
            d.index = moved(r)
            ws.row_dimensions[moved(r)] = d
    if delta > 0 and at - 1 in dict(dims):
        ws.row_dimensions[at].height = dict(dims)[at - 1].height
    for coord in merged:
        rng = CellRange(coord)
        top, bottom = rng.min_row, rng.max_row
        if delta > 0:
            top, bottom = top + (top >= at), bottom + (bottom >= at)  # sisipan di dalam merged: memanjang
        else:
            top, bottom = top - (top > at), bottom - (bottom >= at)
        if bottom > top or (bottom == top and rng.max_col > rng.min_col):
            ws.merge_cells(start_row=top, start_column=rng.min_col, end_row=bottom, end_column=rng.max_col)
        if delta > 0 and rng.min_row == rng.max_row == at - 1:  # merged satu baris di baris template ikut disalin
            ws.merge_cells(start_row=at, start_column=rng.min_col, end_row=at, end_column=rng.max_col)


def change_cols(wb, ws, at, delta, track_last=False, band=None, last_only=()):
    """Padanan change_rows untuk kolom: sisipkan kolom `at` (delta=1; rumus & format disalin dari kolom di kirinya,
    isian kosong, nomor urut di baris judul +1) atau hapus kolom `at` (delta=-1). Hanya baris di dalam `band`
    (baris_atas, baris_bawah; default seluruh sheet) yang digeser, sehingga isi di atas tabel bulan (mis. kotak
    ringkasan durasi) tetap di tempatnya. Rumus semua sheet, rujukan chart, dan merged cell di pita ikut digeser.
    last_only: baris yang selnya hanya ada di kolom bulan terakhir (mis. cek =M47-D46); sel itu ikut pindah ke
    kolom terakhir yang baru."""
    top, bottom = band or (1, ws.max_row)
    rows, last_col = range(top, bottom + 1), ws.max_column
    for rng in ws.merged_cells.ranges:
        if rng.min_row < top <= rng.max_row and rng.max_col >= at - 1:
            raise ValueError(f"merged cell {rng.coord} memotong batas tabel bulan")
    if delta < 0:  # sel khusus bulan terakhir pindah ke kolom sebelumnya (jadi bulan terakhir)
        for r in last_only:
            src, left = ws.cell(r, at), ws.cell(r, at - 1)
            if src.value not in (None, "") and not isinstance(left, MergedCell):
                left.value, left._style = shift_month_refs(src.value, ws.title, at, -1), copy.copy(src._style)
    for w, c, text in formula_cells(wb):
        set_formula(c, shift_formula(text, w.title, ws.title, at, delta, "col", track_last, (top, bottom)))
    for w in wb.worksheets:  # rujukan data chart, mis. '6. Actual Cost'!$E$47:$M$47
        for ch in w._charts:
            for sub in getattr(ch, "_charts", None) or [ch]:
                for s in sub.series:
                    for src in (s.cat, s.val, s.xVal, s.yVal, s.tx):
                        for kind in ("numRef", "strRef"):
                            ref = getattr(src, kind, None) if src is not None else None
                            if ref is not None and ref.f:
                                ref.f = shift_ref(ref.f, ws.title, ws.title, at, delta, "col", False, (top, bottom))
    merged = [CellRange(r.coord) for r in ws.merged_cells.ranges if r.min_row >= top]
    for rng in merged:
        ws.unmerge_cells(rng.coord)
    L = get_column_letter
    if delta > 0:
        if at <= last_col:
            ws.move_range(f"{L(at)}{top}:{L(last_col)}{bottom}", cols=1)
        for r in rows:
            src, dst = ws.cell(r, at - 1), ws.cell(r, at)
            dst._style = copy.copy(src._style)
            if is_formula(src):
                dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
            else:  # nomor bulan (…, 8, 9 -> 10); isian lain dikosongkan
                prev = ws.cell(r, at - 2).value if at > 2 else None
                try:
                    if src.value is not None and float(prev) == float(src.value) - 1:
                        n = int(float(src.value)) + 1
                        dst.value = str(n) if isinstance(src.value, str) else n
                except (TypeError, ValueError):
                    pass
        for r in last_only:  # sel khusus bulan terakhir pindah ke kolom baru (bulan terakhir yang baru)
            src, dst = ws.cell(r, at - 1), ws.cell(r, at)
            if src.value not in (None, ""):
                dst.value = shift_month_refs(src.value, ws.title, at - 1, 1)
                src.value = None
        if ws.column_dimensions[L(at - 1)].width:
            ws.column_dimensions[L(at)].width = ws.column_dimensions[L(at - 1)].width
    else:
        for r in rows:
            ws.cell(r, at).value = None
        if at < last_col:
            ws.move_range(f"{L(at + 1)}{top}:{L(last_col)}{bottom}", cols=-1)
    for rng in merged:
        lo, hi = rng.min_col, rng.max_col
        if delta > 0:
            lo, hi = lo + (lo >= at), hi + (hi >= at)  # sisipan di dalam merged: melebar
        else:
            lo, hi = lo - (lo > at), hi - (hi >= at)
        if hi > lo or (hi == lo and rng.max_row > rng.min_row):
            ws.merge_cells(start_row=rng.min_row, start_column=lo, end_row=rng.max_row, end_column=hi)


def shift_month_refs(value, sheet, last, delta):
    """Rumus sel khusus bulan terakhir dipindah `delta` kolom: rujukan relatif ke kolom bulan terakhir (`last`) di
    sheet yang sama ikut bergeser, rujukan lain tetap (mis. =M47-D46 -> =N47-D46)."""
    if not (isinstance(value, str) and value.startswith("=")):
        return value
    tok = Tokenizer(value)
    for t in tok.items:
        m = CELL_REF.match(t.value) if t.type == "OPERAND" and t.subtype == "RANGE" else None
        if m and not m["r2"] and not m["c1"].startswith("$") \
                and (m["sheet"] is None or m["sheet"][:-1].strip("'").replace("''", "'") == sheet) \
                and column_index_from_string(m["c1"]) == last:
            t.value = f"{m['sheet'] or ''}{get_column_letter(last + delta)}{m['a1']}{m['r1']}"
    return "=" + "".join(t.value for t in tok.items)


def month_axes(book):
    """Sumbu bulan tiap sheet: baris judul bulan (pertama kali dikenali dari >= 3 tanggal berurutan, lalu posisinya
    disimpan di monitoring_books.month_axes agar tetap dikenali walau tinggal 1-2 bulan).
    -> [{title, row, cols, months ('YYYY-MM'), year_row}] ; year_row = baris label tahun di atasnya (atau None)."""
    meta = dict(book.get("axes_meta") or {})
    if any("last_only" not in v for v in meta.values()):  # tersimpan oleh versi sebelumnya: kenali ulang
        meta = {}
    if not meta:  # pertama kali: cari baris berisi >= 3 tanggal berurutan
        for ws in book["wb"].worksheets:
            for r in range(1, min(ws.max_row, 40) + 1):
                dates = [c for c in range(1, ws.max_column + 1) if isinstance(book_value(book, ws.cell(r, c)), datetime)]
                run = []
                for c in dates:
                    run = run + [c] if run and c == run[-1] + 1 else ([c] if len(run) < 3 else run)
                if len(run) >= 3:
                    # baris yang hanya berisi di bulan terakhir (bulan pertama kosong), mis. cek =M47-D46
                    last_only = [rr for rr in range(r + 1, ws.max_row + 1)
                                 if ws.cell(rr, run[-1]).value not in (None, "") and not isinstance(ws.cell(rr, run[-1]), MergedCell)
                                 and ws.cell(rr, run[0]).value in (None, "") and not isinstance(ws.cell(rr, run[0]), MergedCell)]
                    meta[ws.title] = {"row": r, "first": run[0], "last_only": last_only}
                    break
        if meta and book.get("pid"):
            db().execute("UPDATE monitoring_books SET month_axes=? WHERE project_id=?", (json.dumps(meta), book["pid"]))
            db().commit()
            book["axes_meta"] = meta
    axes = []
    for title, pos in meta.items():
        if title not in book["wb"].sheetnames:
            continue
        ws, r, cols, months = book["wb"][title], pos["row"], [], []
        c = pos["first"]
        while c <= ws.max_column and isinstance(v := book_value(book, ws.cell(r, c)), datetime):
            cols.append(c)
            months.append(v.strftime("%Y-%m"))
            c += 1
        if not cols:
            continue
        is_year = lambda rr: any(re.fullmatch(r"(19|20)\d\d", str(book_value(book, ws.cell(rr, c)) or "").strip()) for c in cols)
        year_row = next((rr for rr in range(r - 1, max(r - 4, 0), -1) if is_year(rr)), None)
        axes.append({"title": title, "row": r, "cols": cols, "year_row": year_row, "months": months,
                     "last_only": pos.get("last_only", [])})
    return axes


def contract_months(start, end):
    """Bulan kolom Cost Planning: mulai bulan sesudah bulan Mulai (EOMONTH(Mulai)+1), sebanyak
    DATEDIF(Mulai, Berakhir, "m") bulan (rumus 'Total Durasi' di sheet). -> ['YYYY-MM', ...]"""
    serial = lambda d: (d - datetime(1899, 12, 30)).days
    n = datedif(serial(start), serial(end), "m")
    if not isinstance(n, int):
        return []
    y, mth = start.year + (start.month == 12), start.month % 12 + 1
    out = []
    for _ in range(n):
        out.append(f"{y}-{mth:02d}")
        y, mth = y + (mth == 12), mth % 12 + 1
    return out


def reshape_months(pid, count):
    """Jumlah kolom bulan di setiap sheet bersumbu bulan (Cost Planning, Actual Cost, Rekap) = count: kolom ditambah
    sesudah bulan terakhir (rumus disalin dari bulan terakhir) atau bulan terakhir dihapus. Editan sel dan batas
    tampilan ikut bergeser. Dibatalkan (error) bila ada rumus yang jadi #REF!."""
    con = db()
    book = load_book(pid)
    row = con.execute("SELECT xlsx, extra_cols FROM monitoring_books WHERE project_id=?", (pid,)).fetchone()
    wb = openpyxl.load_workbook(io.BytesIO(row["xlsx"]))
    refs_before = {(w.title, c.coordinate) for w, c, t in formula_cells(wb) if "#REF!" in t}
    limits = {w.title: (mc, mr) for w, mc, mr, _l in book_layout(book)}
    extra = json.loads(row["extra_cols"])
    moves = []  # (sheet, at, delta) untuk menggeser editan
    for ax in month_axes(book):
        ws, n, last = wb[ax["title"]], len(ax["cols"]), ax["cols"][-1]
        band = (ax["year_row"] or ax["row"], ws.max_row)  # baris di atas tabel bulan (judul, kotak ringkasan) tetap
        base = next((column_index_from_string(col) for key, col, _r, _l in MONITORING_LAYOUT if key in ws.title.lower()), 0)
        for i in range(abs(count - n)):
            at, delta = (last + 1 + i, 1) if count > n else (last - i, -1)
            try:
                change_cols(wb, ws, at, delta, track_last=True, band=band, last_only=ax["last_only"])
            except ValueError as e:
                return f"Kolom bulan {ws.title} tidak bisa diubah: {e}."
            moves.append((ws.title, at, delta, band))
            if at <= base + extra.get(ws.title, 0):  # kolom tampilan (mis. Actual s/d N) ikut melebar/menyempit
                extra[ws.title] = extra.get(ws.title, 0) + delta
    if not moves:
        return None
    # Rumus yang jadi #REF! karena kolom bulannya dihapus: tolak bila sel itu tampil di dashboard atau dipakai rumus
    # lain. Sel hitungan bantu yang tersembunyi dan tidak dipakai (mis. coretan di bawah tabel) tidak menghalangi.
    broken = [(w, c, t) for w, c, t in formula_cells(wb) if "#REF!" in t and (w.title, c.coordinate) not in refs_before]
    texts = [(w.title, t) for w, _c, t in formula_cells(wb)]
    def matters(w, c):
        mc, mr = limits.get(w.title, (0, 0))
        if c.column <= mc and c.row <= mr:
            return True
        pat = re.compile(r"(?<![A-Z$])\$?%s\$?%d(?!\d)" % (get_column_letter(c.column), c.row))
        return any(pat.search(t) and (title == w.title or w.title in t) for title, t in texts)
    blocking = [f"{w.title}!{c.coordinate}" for w, c, _t in broken if matters(w, c)]
    if blocking:
        return ("Jumlah bulan tidak bisa dikurangi: rumus di " + ", ".join(blocking[:3])
                + " memakai kolom bulan yang akan dihapus.")
    buf = io.BytesIO()
    wb.save(buf)
    con.execute("UPDATE monitoring_books SET xlsx=?, extra_cols=? WHERE project_id=?", (buf.getvalue(), json.dumps(extra), pid))
    for title, at, delta, band in moves:
        shift_edits(pid, title, at, delta, "col", band)
    con.commit()
    _books.pop(pid, None)
    return None


def shift_edits(pid, title, at, delta, axis, band=None):
    """Geser alamat editan sel sheet `title` setelah baris/kolom `at` disisipkan (delta=1) atau dihapus (-1).
    band (kolom): hanya editan di baris pita yang bergeser."""
    con = db()
    edits = con.execute("SELECT cell, value FROM monitoring_edits WHERE project_id=? AND sheet=?", (pid, title)).fetchall()
    con.execute("DELETE FROM monitoring_edits WHERE project_id=? AND sheet=?", (pid, title))
    for e in edits:
        col, r = re.fullmatch(r"([A-Z]+)(\d+)", e["cell"]).groups()
        i = int(r) if axis == "row" else column_index_from_string(col)
        if not (band and not band[0] <= int(r) <= band[1]):
            i = i + (i >= at) if delta > 0 else (None if i == at else i - (i > at))
        if i:
            cell = f"{col}{i}" if axis == "row" else f"{get_column_letter(i)}{r}"
            con.execute("INSERT INTO monitoring_edits (project_id, sheet, cell, value) VALUES (?,?,?,?)",
                        (pid, title, cell, e["value"]))


def relabel_years(pid):
    """Label tahun di atas judul bulan dihitung dari bulannya: satu label per tahun, di-merge selebar bulannya."""
    book = load_book(pid)
    con = db()
    wb = openpyxl.load_workbook(io.BytesIO(con.execute("SELECT xlsx FROM monitoring_books WHERE project_id=?",
                                                       (pid,)).fetchone()["xlsx"]))
    for ax in month_axes(book):
        if ax["year_row"] is None:
            continue
        ws, yr, cols = wb[ax["title"]], ax["year_row"], ax["cols"]
        style = copy.copy(ws.cell(yr, cols[0])._style)
        for rng in [r for r in ws.merged_cells.ranges if r.min_row == r.max_row == yr and r.min_col >= cols[0] and r.max_col <= cols[-1]]:
            ws.unmerge_cells(rng.coord)
        runs = []
        for c, ym in zip(cols, ax["months"]):
            ws.cell(yr, c).value, ws.cell(yr, c)._style = None, copy.copy(style)
            if runs and runs[-1][0] == ym[:4]:
                runs[-1][2] = c
            else:
                runs.append([ym[:4], c, c])
        for year, lo, hi in runs:
            ws.cell(yr, lo).value = year
            if hi > lo:
                ws.merge_cells(start_row=yr, start_column=lo, end_row=yr, end_column=hi)
    buf = io.BytesIO()
    wb.save(buf)
    con.execute("UPDATE monitoring_books SET xlsx=? WHERE project_id=?", (buf.getvalue(), pid))
    con.commit()
    _books.pop(pid, None)


def month_inputs(book, axes):
    """Isian angka/teks (bukan rumus) di kolom bulan: {(sheet, baris, 'YYYY-MM'): nilai}, juga daftar selnya.
    Baris = blok di bawah judul bulan sampai baris kosong pertama; baris nomor bulan (1, 2, 3, …) dilewati."""
    values, cells = {}, []
    for ax in axes:
        ws = book["wb"][ax["title"]]
        r = ax["row"] + 1
        while r <= ws.max_row and any(ws.cell(r, c).value not in (None, "") for c in range(1, ax["cols"][-1] + 1)):
            raw = [ws.cell(r, c).value for c in ax["cols"]]
            numbering = all(str(v).strip() == str(i + 1) for i, v in enumerate(raw))
            for c, ym in ([] if numbering else zip(ax["cols"], ax["months"])):
                cell = ws.cell(r, c)
                if is_formula(cell):
                    continue
                cells.append((cell, ym))
                v = book["edits"].get((ws.title, cell.coordinate), cell.value)
                if v not in (None, "", 0):
                    values[(ws.title, r, ym)] = v
            r += 1
    return values, cells


def clean_path(raw):
    """Path yang ditempel dari Finder/terminal: buang spasi, tanda kutip, dan escape spasi."""
    return os.path.expanduser(raw.strip().strip("'\"").replace("\\ ", " "))


# Isian proyek manual, mengikuti sheet "1. Data Project": (kunci, label, tipe input)
MANUAL_FIELDS = [
    ("division", "Divisi", "text"), ("project", "Nama pekerjaan", "text"), ("account", "Account project DSI", "text"),
    ("contract_no", "No. kontrak", "text"), ("contract_date", "Tanggal kontrak", "date"),
    ("start_date", "Mulai kontrak", "date"), ("end_date", "Kontrak berakhir", "date"),
    ("add_no", "No. addendum kontrak", "text"), ("add_date", "Tanggal addendum", "date"),
    ("value", "Nilai kontrak (excl. PPN)", "number"), ("value_ppn", "Nilai kontrak (incl. PPN)", "number"),
    ("item", "Item", "text"), ("scope", "Scope of work", "text"), ("down_payment", "Down payment (DP)", "text"),
    ("retention", "Retensi", "text"), ("bank_guarantee", "Bank garansi", "text"),
]


def xlsx_path_error(path):
    if not path:
        return "Path file wajib diisi untuk proyek dari file Excel."
    if not path.lower().endswith((".xlsx", ".xlsm")):
        return "File harus berformat .xlsx atau .xlsm."
    if not os.path.isfile(path):
        return "File tidak ditemukan. Cek lagi path-nya."
    return None


def monitoring_or_404(pid):
    p = db().execute("SELECT * FROM monitoring_projects WHERE id=?", (pid,)).fetchone()
    if p is None:
        abort(404)
    return p


def info_from_form(f):
    """(info dict, pesan error) dari isian MANUAL_FIELDS di form."""
    info = {}
    for key, label, kind in MANUAL_FIELDS:
        value = f.get(key, "").strip()
        if kind == "number" and value:
            try:
                value = float(value)
            except ValueError:
                return None, f"{label} harus angka."
        info[key] = value
    return info, None


# Label di sheet "1. Data Project" -> kunci MANUAL_FIELDS
XLSX_INFO_LABELS = {
    "division": "division", "project": "project", "account project dsi": "account", "no. contract": "contract_no",
    "date contract": "contract_date", "start contract": "start_date", "contract expired": "end_date",
    "add no. contract": "add_no", "date add. contract": "add_date", "nilai kontrak": "value",
    "kontrak awal": "value_ppn", "item": "item", "scope of work": "scope", "down payment (dp)": "down_payment",
    "retensi": "retention", "bank garansi": "bank_guarantee",
}
# Baris di sheet rekap progress: cash-in per bulan, cost rencana & actual kumulatif
XLSX_COST_LABELS = {"realisasi pembayaran": "cash_in", "planning cost": "plan", "actual cost": "actual"}


def sheet_rows(ws):
    """Tiap sel teks: (label, sel-sel sesudahnya di baris yang sama). Satu baris bisa memuat beberapa tabel."""
    for row in ws.iter_rows():
        for i, c in enumerate(row):
            if isinstance(c.value, str) and c.value.strip():
                yield c.value.strip().lower(), row[i + 1:]


def data_project_cells(ws):
    """Sheet Data Project -> {kunci MANUAL_FIELDS: sel nilai}. Baris tanpa nilai memakai kolom nilai terbanyak."""
    found, rows = {}, {}
    for label, rest in sheet_rows(ws):
        key = XLSX_INFO_LABELS.get(label)
        if key and key not in rows and rest:
            rows[key] = rest[0].row
            found[key] = next((c for c in rest if c.value is not None and str(c.value).strip() != ""), None)
    cols = [c.column for c in found.values() if c is not None]
    col = max(set(cols), key=cols.count) if cols else None
    return {k: c if c is not None else ws.cell(rows[k], col) for k, c in found.items() if c is not None or col}


def read_monitoring_xlsx(wb, value=lambda c: c.value):
    """Data proyek + cost per bulan dari workbook format Detail Monitoring Project: (info, costs).
    value(cell) memberi nilai sel (nilai cache Excel, atau hasil hitung ulang pycel)."""
    sheets = [ws for ws in wb.worksheets if not ws.title.lower().startswith("arsip")]
    info = {}
    data_ws = next((ws for ws in sheets if "data project" in ws.title.lower()), None)
    for key, cell in (data_project_cells(data_ws) if data_ws else {}).items():
        v = value(cell)
        if v is None:
            v = ""
        elif isinstance(v, datetime):
            v = v.date().isoformat()
        elif isinstance(v, str):
            v = v.strip()
            v = "" if v == "-" and key in ("contract_no", "add_no") else v
        info[key] = v if MANUAL_FIELD_KIND[key] != "number" or isinstance(v, (int, float)) else ""
    costs = []
    for ws in sheets:
        found = {XLSX_COST_LABELS[label]: rest for label, rest in sheet_rows(ws) if label in XLSX_COST_LABELS}
        if len(found) < len(XLSX_COST_LABELS):
            continue
        # bulan: sel bertipe tanggal di kolom yang sama, pada baris di atas baris data
        first = found["plan"][0]
        months = {}
        for row in ws.iter_rows(max_row=first.row - 1, min_col=first.column):
            for c in row:
                v = value(c)
                if isinstance(v, datetime):
                    months[c.column] = v.strftime("%Y-%m")
        num = lambda c: v if isinstance(v := value(c), (int, float)) and not isinstance(v, bool) else 0
        prev = {"plan": 0, "actual": 0}
        for col in sorted(months):
            by = {k: next((num(c) for c in rest if c.column == col), 0) for k, rest in found.items()}
            costs.append({"month": months[col], "cash_in": by["cash_in"],
                          "plan": by["plan"] - prev["plan"], "actual": by["actual"] - prev["actual"]})
            prev = {"plan": by["plan"], "actual": by["actual"]}
        break
    return info, costs


MANUAL_FIELD_KIND = {k: kind for k, _label, kind in MANUAL_FIELDS}


MONITORING_PER_PAGE = 10


def monitoring_list(**form_state):
    con = db()
    total = con.execute("SELECT COUNT(*) FROM monitoring_projects").fetchone()[0]
    pages = max(1, math.ceil(total / MONITORING_PER_PAGE))
    page = min(max(request.args.get("page", 1, type=int), 1), pages)
    rows = [{**dict(p), "file_name": os.path.basename(p["xlsx_path"])}
            for p in con.execute("SELECT p.*, b.imported_at FROM monitoring_projects p"
                                 " LEFT JOIN monitoring_books b ON b.project_id = p.id ORDER BY p.name LIMIT ? OFFSET ?",
                                 (MONITORING_PER_PAGE, (page - 1) * MONITORING_PER_PAGE))]
    form_state.update(page=page, pages=pages, total=total, first=(page - 1) * MONITORING_PER_PAGE + 1)
    form_state.setdefault("pre", {})
    form_state.setdefault("pre_costs", [])
    open_form = bool(form_state.get("error") or form_state.get("xlsx_error") or form_state.get("xlsx_read")
                     or form_state.get("link_error"))
    return render_template("monitoring.html", active="monitoring", monitoring_pid=None, rows=rows,
                           fields=MANUAL_FIELDS, open_form=open_form, **form_state)


@app.route("/monitoring", methods=["GET", "POST"])
def monitoring():
    if request.method == "GET":
        return monitoring_list()
    f = request.form
    name = f.get("name", "").strip()
    info, error = info_from_form(f)
    try:
        costs = json.loads(f.get("costs_json") or "[]")
        costs = [(c["month"], float(c["plan"]), float(c["actual"]), float(c["cash_in"])) for c in costs]
    except (ValueError, KeyError, TypeError):
        costs, error = [], "Data cost dari file Excel rusak. Baca ulang file-nya."
    if not name:
        error = "Nama proyek wajib diisi."
    if error:
        return monitoring_list(error=error, pre={**(info or {}), "name": name}, pre_costs=[])
    con = db()
    pid = con.execute("INSERT INTO monitoring_projects (name, xlsx_path, source, info) VALUES (?,?,?,?)",
                      (name, "", "manual", json.dumps(info))).lastrowid
    con.executemany("INSERT OR REPLACE INTO monitoring_costs (project_id, month, plan, actual, cash_in) VALUES (?,?,?,?,?)",
                    [(pid, *c) for c in costs])
    con.commit()
    return redirect(url_for("monitoring_project", pid=pid))


@app.post("/monitoring/link")
def monitoring_link():
    """Impor file Excel lewat path: sheet monitoring disimpan di database, sel input bisa diedit di dashboard."""
    path = clean_path(request.form.get("xlsx_path", ""))
    error = xlsx_path_error(path)
    if error:
        return monitoring_list(link_error=error, link_path=path)
    name = os.path.splitext(os.path.basename(path))[0]
    try:  # nama dari sheet Data Project bila ada; bisa diganti lewat Edit proyek
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            info, _costs = read_monitoring_xlsx(openpyxl.load_workbook(path, data_only=True))
        name = " - ".join(str(info[k]) for k in ("account", "project") if info.get(k)) or name
        snapshot = make_snapshot(path)
    except Exception as e:
        return monitoring_list(link_error=f"Gagal membaca file Excel: {e}", link_path=path)
    con = db()
    pid = con.execute("INSERT INTO monitoring_projects (name, xlsx_path, source) VALUES (?,?,?)",
                      (name, path, "xlsx")).lastrowid
    con.execute("INSERT INTO monitoring_books (project_id, xlsx, imported_at) VALUES (?,?,?)",
                (pid, snapshot, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    con.commit()
    return redirect(url_for("monitoring_project", pid=pid))


@app.post("/monitoring/read-xlsx")
def monitoring_read_xlsx():
    file = request.files.get("xlsx")
    if not file or not file.filename.lower().endswith((".xlsx", ".xlsm")):
        return monitoring_list(xlsx_error="Pilih file .xlsx.")
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(file.stream, data_only=True)
        info, costs = read_monitoring_xlsx(wb)
    except Exception as e:  # bukan file Excel yang valid
        return monitoring_list(xlsx_error=f"Gagal membaca file Excel: {e}")
    if not info and not costs:
        return monitoring_list(xlsx_error="Format file tidak dikenali: sheet Data Project dan rekap cost tidak ditemukan.")
    name = " - ".join(str(info[k]) for k in ("account", "project") if info.get(k))
    return monitoring_list(pre={**info, "name": name}, pre_costs=costs, xlsx_read=True,
                           xlsx_name=file.filename)


def cost_summary(info, rows):
    """Penjumlahan otomatis dari cost per bulan ({month, plan, actual, cash_in}, urut bulan):
    (baris + kumulatif & deviasi, total, ringkasan keuangan terhadap nilai kontrak)."""
    costs, total = [], {"plan": 0.0, "actual": 0.0, "cash_in": 0.0}
    for c in rows:
        for k in total:
            total[k] += c[k]
        costs.append({**c, "cum_plan": total["plan"], "cum_actual": total["actual"],
                      "cum_cash_in": total["cash_in"], "deviation": total["plan"] - total["actual"]})
    value = info.get("value") if isinstance(info.get("value"), (int, float)) else 0
    summary = {"value": value, **total, "deviation": total["plan"] - total["actual"],
               "remaining": value - total["cash_in"], "margin": value - total["actual"],
               "actual_pct": total["actual"] / value * 100 if value else None}
    return costs, total, summary


@app.route("/monitoring/<int:pid>")
@app.route("/monitoring/<int:pid>/<int:sheet_idx>")
def monitoring_project(pid, sheet_idx=0):
    p = monitoring_or_404(pid)
    common = {"active": "monitoring", "monitoring_pid": pid, "p": p, "saved": request.args.get("saved") == "1",
              "edit_error": request.args.get("error", "")[:200], "open_edit": request.args.get("edit") == "1"}
    if p["source"] == "manual":
        info = json.loads(p["info"] or "{}")
        rows = {r["month"]: {**dict(r), "pum": 0} for r in
                db().execute("SELECT * FROM monitoring_costs WHERE project_id=?", (pid,))}
        for (_cat, ym), amt in pum_actuals(pid).items():  # PUM masuk ke cost actual bulan tanggal PUM
            row = rows.setdefault(ym, {"id": None, "month": ym, "plan": 0, "actual": 0, "cash_in": 0, "pum": 0})
            row["actual"] += amt
            row["pum"] += amt
        costs, total, summary = cost_summary(info, [rows[m] for m in sorted(rows)])
        chart = {"title": "Cost monitoring", "labels": [month_label(c["month"]) for c in costs], "series": [
            {"label": "Cost actual per bulan", "values": [c["actual"] for c in costs], "bar": True},
            {"label": "Kumulatif rencana", "values": [c["cum_plan"] for c in costs], "bar": False},
            {"label": "Kumulatif actual", "values": [c["cum_actual"] for c in costs], "bar": False},
            {"label": "Kumulatif cash-in", "values": [c["cum_cash_in"] for c in costs], "bar": False},
        ]}
        return render_template("monitoring_manual.html", fields=MANUAL_FIELDS, info=info, costs=costs,
                               total=total, summary=summary, charts=[chart] if costs else [], **common)
    book, error = monitoring_book(pid, p)
    layout = book_layout(book) if book else []
    if layout and not 0 <= sheet_idx < len(layout):
        abort(404)
    sheet = charts = None
    if layout:
        ws, max_col, max_row, is_list = layout[sheet_idx]
        value = lambda c: book_value(book, c)
        edited = {cell for (s, cell) in book["edits"] if s == ws.title}
        limits = {w.title: (mc, mr) for w, mc, mr, _l in layout}
        inside = lambda c: c.parent.title in limits and c.column <= limits[c.parent.title][0] \
            and c.row <= limits[c.parent.title][1]
        pum = {cell: amt for (s, cell), amt in book["pum_extra"].items() if s == ws.title}
        rows, widths = read_sheet(ws, value, max_col, max_row, edited, pum)
        sheet = {"title": ws.title, "rows": rows, "widths": widths,
                 "charts": read_charts(ws, book["wb"], value, inside), "images": len(ws._images), "is_list": is_list}
    # Ringkasan dari Data Project + rekap cost, memakai nilai setelah editan dan hitung ulang rumus
    info, costs = read_monitoring_xlsx(book["wb"], lambda c: book_value(book, c)) if book else ({}, [])
    has_finance = bool(info.get("value") or costs)
    _costs, total, summary = cost_summary(info, costs) if has_finance else (None, None, None)
    data_ws = find_sheet(book["wb"], "data project") if book else None
    cells = data_project_cells(data_ws) if data_ws else {}
    info_fields = [(k, label, kind, is_formula(cells[k])) for k, label, kind in MANUAL_FIELDS if k in cells]
    return render_template("monitoring_project.html", sheets=[w.title for w, _c, _r, _l in layout], sheet_idx=sheet_idx,
                           error=error, sheet=sheet, total=total, summary=summary, info=info, info_fields=info_fields,
                           n_edits=len(book["edits"]) if book else 0, imported_at=book and book["imported_at"],
                           file_name=os.path.basename(p["xlsx_path"]), **common)


def monitoring_book(pid, p):
    """(book, error). Proyek dari versi lama (belum punya snapshot) diimpor sekali dari path-nya."""
    try:
        book = load_book(pid)
        if book is None:
            save_snapshot(pid, p["xlsx_path"])
            book = load_book(pid)
        return book, None
    except Exception as e:
        return None, f"Gagal memuat isi Excel: {e}"


def calc(expr):
    """Hitung ekspresi angka + - * / dan kurung; selain itu None (tetap teks)."""
    if not re.fullmatch(r"[\d\s.+\-*/()]+", expr):
        return None
    ops = {ast.Add: lambda a, b: a + b, ast.Sub: lambda a, b: a - b,
           ast.Mult: lambda a, b: a * b, ast.Div: lambda a, b: a / b}

    def ev(node):
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return node.value
        if isinstance(node, ast.BinOp) and type(node.op) in ops:
            return ops[type(node.op)](ev(node.left), ev(node.right))
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.USub, ast.UAdd)):
            return -ev(node.operand) if isinstance(node.op, ast.USub) else ev(node.operand)
        raise ValueError

    try:
        return float(ev(ast.parse(expr, mode="eval").body))
    except (SyntaxError, ValueError, ZeroDivisionError):
        return None


def parse_cell_input(text, cell):
    """Isian form -> nilai sel: kosong = None, tanggal (sel berformat tanggal) = serial Excel,
    angka (1.234.567,89 / 1234567.89 / 12%) = float, selain itu teks."""
    s = text.strip()
    if not s:
        return None
    if s.startswith("="):  # hitungan sederhana ala Excel, mis. =900000+1350000 -> 2250000
        n = calc(s[1:])
        return s if n is None else n
    if is_date_format(cell.number_format) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return to_excel(datetime.strptime(s, "%Y-%m-%d"))
    t = s.replace("Rp", "").replace(" ", "")
    pct = t.endswith("%")
    t = t.rstrip("%")
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", t):  # 1.500.000 = pemisah ribuan
        t = t.replace(".", "")
    try:
        n = float(t)
    except ValueError:
        return s
    return n / 100 if pct else n


def save_cell_edit(pid, cell, new):
    """Simpan editan satu sel input; nilai yang sama dengan isi asli menghapus editannya."""
    con = db()
    original = cell.value
    if isinstance(original, datetime):
        original = to_excel(original)  # editan tanggal disimpan sebagai serial Excel
    same = new == original or (new is None and original in (None, "")) or \
        (isinstance(new, str) and isinstance(original, str) and new.strip() == original.strip())
    if same:
        con.execute("DELETE FROM monitoring_edits WHERE project_id=? AND sheet=? AND cell=?",
                    (pid, cell.parent.title, cell.coordinate))
    else:
        con.execute("INSERT OR REPLACE INTO monitoring_edits (project_id, sheet, cell, value) VALUES (?,?,?,?)",
                    (pid, cell.parent.title, cell.coordinate, json.dumps(new)))


@app.post("/monitoring/<int:pid>/<int:sheet_idx>/cell")
def monitoring_cell(pid, sheet_idx):
    """Edit satu sel input di sheet yang tampil. reset=1 mengembalikan nilai asli dari file."""
    p = monitoring_or_404(pid)
    book, error = monitoring_book(pid, p)
    layout = book_layout(book) if book else []
    if not 0 <= sheet_idx < len(layout):
        abort(404)
    ws, max_col, max_row, _is_list = layout[sheet_idx]
    ref = request.form.get("cell", "").upper()
    back = lambda **kw: redirect(url_for("monitoring_project", pid=pid, sheet_idx=sheet_idx, **kw) + f"#c-{ref}")
    m = re.fullmatch(r"([A-Z]{1,3})([0-9]{1,6})", ref)
    if not m or column_index_from_string(m.group(1)) > max_col or int(m.group(2)) > max_row:
        return back(error="Sel di luar area yang bisa diedit.")
    cell = ws[ref]
    if is_formula(cell):
        return back(error=f"Sel {ref} berisi rumus dan dihitung otomatis, tidak bisa diedit.")
    new = cell.value if request.form.get("reset") == "1" else parse_cell_input(request.form.get("value", ""), cell)
    save_cell_edit(pid, cell, new)
    db().commit()
    return back(saved=1)


@app.post("/monitoring/<int:pid>/<int:sheet_idx>/rows")
def monitoring_rows(pid, sheet_idx):
    """Sheet daftar: action=insert menyisipkan baris di bawah `row`, action=delete menghapus baris `row`."""
    p = monitoring_or_404(pid)
    book, error = monitoring_book(pid, p)
    layout = book_layout(book) if book else []
    if not 0 <= sheet_idx < len(layout) or not layout[sheet_idx][3]:
        abort(404)
    title, max_row = layout[sheet_idx][0].title, layout[sheet_idx][2]
    action, row = request.form.get("action"), request.form.get("row", type=int)
    back = lambda r, **kw: redirect(url_for("monitoring_project", pid=pid, sheet_idx=sheet_idx, **kw) + f"#c-A{r}")
    if action not in ("insert", "delete") or not row or not 1 <= row <= max_row:
        return back(row or 1, error="Pilih baris di dalam area sheet.")
    con = db()
    wb = openpyxl.load_workbook(io.BytesIO(
        con.execute("SELECT xlsx FROM monitoring_books WHERE project_id=?", (pid,)).fetchone()["xlsx"]))
    ws = wb[title]
    if action == "delete":
        refs = row_references(wb, ws, row)
        if refs:
            more = f" dan {len(refs) - 3} lainnya" if len(refs) > 3 else ""
            return back(row, error=f"Baris {row} tidak bisa dihapus: dipakai rumus di {', '.join(refs[:3])}{more}.")
    at, delta = (row + 1, 1) if action == "insert" else (row, -1)
    change_rows(wb, ws, at, delta)
    buf = io.BytesIO()
    wb.save(buf)
    extra = json.loads(con.execute("SELECT extra_rows FROM monitoring_books WHERE project_id=?", (pid,)).fetchone()[0])
    extra[title] = extra.get(title, 0) + delta
    con.execute("UPDATE monitoring_books SET xlsx=?, extra_rows=? WHERE project_id=?", (buf.getvalue(), json.dumps(extra), pid))
    # editan sel di sheet ini ikut pindah baris; editan di baris yang dihapus dibuang
    edits = con.execute("SELECT cell, value FROM monitoring_edits WHERE project_id=? AND sheet=?", (pid, title)).fetchall()
    con.execute("DELETE FROM monitoring_edits WHERE project_id=? AND sheet=?", (pid, title))
    for e in edits:
        col, r = re.fullmatch(r"([A-Z]+)(\d+)", e["cell"]).groups()
        r = int(r)
        r = r + (r >= at) if delta > 0 else (None if r == at else r - (r > at))
        if r:
            con.execute("INSERT INTO monitoring_edits (project_id, sheet, cell, value) VALUES (?,?,?,?)",
                        (pid, title, f"{col}{r}", e["value"]))
    con.commit()
    _books.pop(pid, None)  # snapshot berubah: workbook & compiler dimuat ulang
    return back(at if delta > 0 else max(row - 1, 1), saved=1)


@app.post("/monitoring/<int:pid>/reimport")
def monitoring_reimport(pid):
    """Ambil ulang isi sheet dari file Excel di path proyek; semua editan sel dibuang."""
    p = monitoring_or_404(pid)
    error = xlsx_path_error(p["xlsx_path"])
    if not error:
        try:
            save_snapshot(pid, p["xlsx_path"])
        except Exception as e:
            error = f"Gagal membaca file Excel: {e}"
    if error:
        return redirect(url_for("monitoring_project", pid=pid, error=error))
    return redirect(url_for("monitoring_project", pid=pid, saved=1))

@app.post("/monitoring/<int:pid>/edit")
def monitoring_edit(pid):
    p = monitoring_or_404(pid)
    f = request.form
    name = f.get("name", "").strip()
    if not name:
        return redirect(url_for("monitoring_project", pid=pid, edit=1, error="Nama proyek wajib diisi."))
    con = db()
    if p["source"] == "manual":
        info, error = info_from_form(f)
        if error:
            return redirect(url_for("monitoring_project", pid=pid, edit=1, error=error))
        con.execute("UPDATE monitoring_projects SET name=?, info=? WHERE id=?", (name, json.dumps(info), pid))
    else:
        path = clean_path(f.get("xlsx_path", ""))
        error = xlsx_path_error(path)
        if error:
            return redirect(url_for("monitoring_project", pid=pid, edit=1, error=error))
        info, error = info_from_form(f)
        book, load_error = monitoring_book(pid, p)
        if error or load_error:
            return redirect(url_for("monitoring_project", pid=pid, edit=1, error=error or load_error))
        data_ws = find_sheet(book["wb"], "data project")
        cells = data_project_cells(data_ws) if data_ws else {}
        # Mulai / berakhir kontrak berubah -> kolom bulan, label tahun, dan isian actual per bulan ikut menyesuaikan
        new_months, moved = None, {}
        if info.get("start_date") and info.get("end_date") and "start_date" in cells and "end_date" in cells:
            old = [book_value(book, cells[k]) for k in ("start_date", "end_date")]
            new = [datetime.strptime(info[k], "%Y-%m-%d") for k in ("start_date", "end_date")]
            if [d.date() if isinstance(d, datetime) else d for d in old] != [d.date() for d in new]:
                new_months = contract_months(*new)
                if len(new_months) < 2:  # 1 kolom = bulan pertama sekaligus terakhir; rumus khusus keduanya tidak bisa dibedakan
                    return redirect(url_for("monitoring_project", pid=pid, edit=1, error=(
                        f"Durasi kontrak minimal 2 bulan (Total Durasi sekarang {len(new_months)} bulan).")))
                if len(new_months) > 60:
                    return redirect(url_for("monitoring_project", pid=pid, edit=1, error="Durasi kontrak maksimal 60 bulan."))
                moved, _cells = month_inputs(book, month_axes(book))
                lost = sorted({(t, ym) for t, _r, ym in moved if ym not in new_months})
                if lost:
                    where = ", ".join(f"{t.split('. ', 1)[-1]} {month_label(ym)}" for t, ym in lost[:4])
                    return redirect(url_for("monitoring_project", pid=pid, edit=1, error=(
                        f"Periode baru tidak mencakup bulan yang sudah berisi angka ({where}). "
                        "Kosongkan angka di bulan itu dulu, atau pilih periode yang mencakupnya.")))
                error = reshape_months(pid, len(new_months))
                if error:
                    return redirect(url_for("monitoring_project", pid=pid, edit=1, error=error))
        for key, cell in cells.items():
            if key not in f:
                continue
            if is_formula(cell):
                continue  # nilai hasil rumus (mis. kontrak incl. PPN) tetap dihitung
            v = info[key]
            if MANUAL_FIELD_KIND[key] == "date" and v:
                v = to_excel(datetime.strptime(v, "%Y-%m-%d"))
            save_cell_edit(pid, cell, v if v != "" else None)
        con.execute("UPDATE monitoring_projects SET name=?, xlsx_path=? WHERE id=?", (name, path, pid))
        if new_months:
            con.commit()
            relabel_years(pid)
            book = load_book(pid)
            _values, month_cells = month_inputs(book, month_axes(book))
            for cell, ym in month_cells:  # isian actual pindah ke kolom bulannya yang baru
                want = moved.get((cell.parent.title, cell.row, ym))
                have = book["edits"].get((cell.parent.title, cell.coordinate), cell.value)
                if want is not None and want != have:
                    save_cell_edit(pid, cell, want)
                elif want is None and have not in (None, "", 0):
                    save_cell_edit(pid, cell, None)
    con.commit()
    return redirect(url_for("monitoring_project", pid=pid, saved=1))


@app.post("/monitoring/<int:pid>/delete")
def monitoring_delete(pid):
    """Hapus dari daftar beserta cost bulanan / snapshot Excel dan editannya; file Excel tidak disentuh."""
    con = db()
    for table in ("monitoring_costs", "monitoring_books", "monitoring_edits", "pum_docs"):
        con.execute(f"DELETE FROM {table} WHERE project_id=?", (pid,))
    con.execute("DELETE FROM monitoring_projects WHERE id=?", (pid,))
    _books.pop(pid, None)
    con.commit()
    return redirect(url_for("monitoring"))


@app.post("/monitoring/<int:pid>/costs")
def monitoring_cost_save(pid):
    """Tambah cost bulanan; bulan yang sudah ada diperbarui (dipakai juga untuk edit)."""
    if monitoring_or_404(pid)["source"] != "manual":
        abort(400)
    f = request.form
    month = f.get("month", "")
    try:
        datetime.strptime(month, "%Y-%m")
        values = [float(f.get(k) or 0) for k in ("plan", "actual", "cash_in")]
    except ValueError:
        return redirect(url_for("monitoring_project", pid=pid, error="Isi bulan dan nilai cost dengan angka."))
    if any(v < 0 for v in values):
        return redirect(url_for("monitoring_project", pid=pid, error="Nilai cost tidak boleh negatif."))
    con = db()
    con.execute("INSERT INTO monitoring_costs (project_id, month, plan, actual, cash_in) VALUES (?,?,?,?,?)"
                " ON CONFLICT(project_id, month) DO UPDATE SET plan=excluded.plan, actual=excluded.actual,"
                " cash_in=excluded.cash_in", (pid, month, *values))
    con.commit()
    return redirect(url_for("monitoring_project", pid=pid, saved=1))


@app.post("/monitoring/<int:pid>/costs/<int:cost_id>/delete")
def monitoring_cost_delete(pid, cost_id):
    db().execute("DELETE FROM monitoring_costs WHERE id=? AND project_id=?", (cost_id, pid))
    db().commit()
    return redirect(url_for("monitoring_project", pid=pid))



# ---------- PUM (Permintaan Uang Muka) per proyek monitoring ----------
# Format dari sheet PUM (sheet pertama yang tampil), dibatasi A1:G58. Tiap PUM = salinan format yang diisi.
# Total per kategori ditambahkan ke Actual Cost proyek pada bulan tanggal PUM.

PUM_MAX_COL, PUM_MAX_ROW, PUM_VALUE_COL = 7, 58, 3  # area format A1:G58; nilai isian kepala di kolom C
PUM_FIELDS = {"tanggal / date": "date", "periode": "periode", "no. pum": "no_pum", "untuk keperluan": "purpose"}
_pums = {}  # pum id -> {key: updated_at, wb, xl}


def pum_layout(ws):
    """Struktur format PUM: sel isian kepala, kolom PIC, kolom keterangan, baris kategori, sel grand total."""
    fields = {key: ws.cell(rest[0].row, PUM_VALUE_COL) for label, rest in sheet_rows(ws)
              if (key := PUM_FIELDS.get(label)) and rest}
    header = next(r for r in range(1, ws.max_row + 1)
                  if str(ws.cell(r, 1).value or "").strip().lower().startswith("no")
                  and str(ws.cell(r, 2).value or "").strip().lower() == "item")
    pic_cols = [c for c in range(PUM_VALUE_COL, ws.max_column + 1) if ws.cell(header + 1, c).value]
    note_col = next((c for c in range(1, ws.max_column + 1)
                     if str(ws.cell(header, c).value or "").strip().lower() == "keterangan"), None)
    items, r = [], header + 2
    while r <= ws.max_row and str(ws.cell(r, 2).value or "").strip().lower() != "total":
        if ws.cell(r, 2).value:
            items.append((r, str(ws.cell(r, 2).value).strip()))
        r += 1
    grand = next((rest[0] for label, rest in sheet_rows(ws) if label == "grand total" and rest), None)
    return {"fields": fields, "pic_cols": pic_cols, "note_col": note_col, "items": items, "grand": grand}


def set_pic_columns(ws, names):
    """Kolom PIC (pengendali UM) PUM = names: kolom ditambah/dikurangi di ujung kanan kolom PIC, kolom sesudahnya
    (Keterangan, tanda tangan) bergeser. Header 'Ammount', judul, TOTAL per kolom, dan GRAND TOTAL disesuaikan."""
    layout = pum_layout(ws)
    first, last = layout["pic_cols"][0], layout["pic_cols"][-1]
    delta = len(names) - (last - first + 1)
    widths = {c: ws.column_dimensions[get_column_letter(c)].width for c in range(1, ws.max_column + 1)}
    merged = [CellRange(r.coord) for r in ws.merged_cells.ranges]
    for rng in merged:
        ws.unmerge_cells(rng.coord)
    if delta > 0:
        ws.insert_cols(last + 1, delta)
        for c in range(last + 1, last + 1 + delta):
            for r in range(1, ws.max_row + 1):
                ws.cell(r, c)._style = copy.copy(ws.cell(r, last)._style)
    elif delta < 0:
        ws.delete_cols(last + delta + 1, -delta)
    new_last = last + delta
    for c, w in widths.items():  # openpyxl tidak menggeser lebar kolom
        target = c if c <= new_last and c <= last else (c + delta if c > last else None)
        if target and w:
            ws.column_dimensions[get_column_letter(target)].width = w
    for c in range(last + 1, new_last + 1):
        ws.column_dimensions[get_column_letter(c)].width = widths.get(last)
    for rng in merged:
        lo, hi = rng.min_col, rng.max_col
        if lo == first and hi == last:  # header Ammount, GRAND TOTAL: selebar kolom PIC
            hi = new_last
        elif lo > last:
            lo, hi = lo + delta, hi + delta
        elif lo <= first and hi >= last:  # judul selebar tabel
            hi += delta
        elif hi > new_last:
            continue
        if hi > lo or rng.max_row > rng.min_row:
            ws.merge_cells(start_row=rng.min_row, start_column=lo, end_row=rng.max_row, end_column=hi)
    layout = pum_layout(ws)
    header = layout["items"][0][0] - 1  # baris nama PIC
    for i, name in enumerate(names):
        ws.cell(header, first + i).value = name
    total_row = layout["items"][-1][0] + 1
    top, bottom = layout["items"][0][0], layout["items"][-1][0]
    for c in range(first, new_last + 1):
        col = get_column_letter(c)
        ws.cell(total_row, c).value = f"=SUM({col}{top}:{col}{bottom})"
    if layout["grand"] is not None:
        layout["grand"].value = "=" + "+".join(f"{get_column_letter(c)}{total_row}" for c in range(first, new_last + 1))


ROMAN_MONTHS = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]


def project_code(p):
    """(nomor proyek, nama proyek) untuk No. PUM: nomor dari Data Project (account) atau '(026-113)' di nama."""
    account = ""
    if p["source"] == "manual":
        account = str(json.loads(p["info"] or "{}").get("account") or "")
    else:
        book, _error = monitoring_book(p["id"], p)
        ws = find_sheet(book["wb"], "data project") if book else None
        cell = data_project_cells(ws).get("account") if ws else None
        account = str(book_value(book, cell) or "") if cell is not None else ""
    if not account:
        m = re.search(r"\(([^)]+)\)", p["name"])
        account = m.group(1) if m else "-"
    return account.strip(), re.sub(r"\s*\([^)]*\)", "", p["name"]).strip().upper()


def pum_number(seq, p, when):
    """001/PUM/026-113/NPEA BRANTAS/XI/2026: urut per proyek per tahun, bulan romawi & tahun dari tanggal PUM."""
    account, name = project_code(p)
    return f"{seq:03d}/PUM/{account}/{name}/{ROMAN_MONTHS[when.month - 1]}/{when.year}"


def pum_controllers():
    """Daftar pengendali UM, urut saat ditambahkan."""
    return db().execute("SELECT id, name FROM pum_controllers ORDER BY id").fetchall()


def load_pum(doc):
    """Baris pum_docs -> {wb, xl} (cache per updated_at)."""
    cached = _pums.get(doc["id"])
    if not cached or cached["key"] != doc["updated_at"]:
        wb = openpyxl.load_workbook(io.BytesIO(doc["xlsx"]))
        polish_pum(wb.active)  # hanya tampilan (tidak disimpan), jadi berlaku juga untuk PUM lama
        cached = _pums[doc["id"]] = {"key": doc["updated_at"], "wb": wb, "xl": compile_workbook(wb)}
    return cached


def polish_pum(ws):
    """Rapikan tampilan PUM (aman diulang): header No./Item/Keterangan di-merge 2 baris (baris judul + baris nama PIC),
    tanda tangan di bawah GRAND TOTAL tanpa kotak dan rata tengah."""
    layout = pum_layout(ws)
    header = layout["items"][0][0] - 2  # baris "No. | Item | Ammount | Keterangan"; baris berikutnya nama PIC
    for col in (1, 2, layout["note_col"]):
        if col and not any(ws.cell(header, col).coordinate in rng for rng in ws.merged_cells.ranges):
            ws.merge_cells(start_row=header, start_column=col, end_row=header + 1, end_column=col)
            top = ws.cell(header, col)
            top.alignment = Alignment(horizontal=top.alignment.horizontal or "center", vertical="center")
    if layout["grand"] is not None:
        for row in ws.iter_rows(min_row=layout["grand"].row + 1):
            for cell in row:
                cell.border = Border()
                if cell.value not in (None, ""):
                    cell.alignment = Alignment(horizontal="center", vertical=cell.alignment.vertical)


def save_pum(doc_id, wb):
    """Simpan isi PUM + ringkasannya (tanggal, no., periode, total, total per kategori) untuk daftar & Actual Cost.
    No. PUM ditulis ulang dari nomor urut + tanggal PUM (bulan romawi & tahun mengikuti tanggal)."""
    doc = db().execute("SELECT d.seq, d.signatures, p.* FROM pum_docs d JOIN monitoring_projects p ON p.id = d.project_id"
                       " WHERE d.id=?", (doc_id,)).fetchone()
    ws = wb.active
    layout = pum_layout(ws)
    raw = layout["fields"]["date"].value if "date" in layout["fields"] else None
    when = from_excel(raw) if isinstance(raw, (int, float)) else raw
    if doc["seq"] and "no_pum" in layout["fields"]:
        layout["fields"]["no_pum"].value = pum_number(doc["seq"], doc, when if isinstance(when, datetime) else datetime.now())
    xl = compile_workbook(wb)
    value = lambda c: book_value({"xl": xl}, c)
    num = lambda v: v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0
    categories = {}
    for r, name in layout["items"]:
        amount = sum(num(value(ws.cell(r, c))) for c in layout["pic_cols"])
        if amount:
            categories[name] = categories.get(name, 0) + amount
    field = lambda k: value(layout["fields"][k]) if k in layout["fields"] else None
    doc_date = field("date")
    buf = io.BytesIO()
    wb.save(buf)
    # tanda tangan hanya berlaku untuk nama yang menandatangani: nama di PUM berubah -> slot itu dilepas
    names = {x["key"]: x["name"].lower() for x in pum_sign_slots({"wb": wb, "xl": xl})}
    signed = {k: v for k, v in json.loads(doc["signatures"] or "{}").items() if names.get(k) == v.lower()}
    db().execute("UPDATE pum_docs SET xlsx=?, updated_at=?, doc_date=?, no_pum=?, periode=?, total=?, categories=?,"
                 " signatures=? WHERE id=?", (buf.getvalue(), datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f"),
                                 doc_date.date().isoformat() if isinstance(doc_date, datetime) else None,
                                 str(field("no_pum") or ""), str(field("periode") or ""),
                                 num(value(layout["grand"])) if layout["grand"] else sum(categories.values()),
                                 json.dumps(categories), json.dumps(signed), doc_id))
    db().commit()


def pum_actuals(pid):
    """{(kategori, 'YYYY-MM'): Rp} dari semua PUM proyek yang bertanggal."""
    out = {}
    for d in db().execute("SELECT doc_date, categories FROM pum_docs WHERE project_id=? AND doc_date IS NOT NULL", (pid,)):
        for cat, amt in json.loads(d["categories"]).items():
            out[(cat, d["doc_date"][:7])] = out.get((cat, d["doc_date"][:7]), 0) + amt
    return out


def pum_doc_or_404(pid, doc_id):
    doc = db().execute("SELECT * FROM pum_docs WHERE id=? AND project_id=?", (doc_id, pid)).fetchone()
    if doc is None:
        abort(404)
    return doc


@app.route("/monitoring/<int:pid>/pum", methods=["GET", "POST"])
def pum_list(pid):
    p = monitoring_or_404(pid)
    con = db()
    template = con.execute("SELECT source, imported_at, xlsx FROM pum_template WHERE id=1").fetchone()
    if request.method == "POST":  # PUM baru dari format: tanggal hari ini, keperluan = nama proyek
        if template is None:
            abort(400)
        names = [c["name"] for c in pum_controllers()]
        if not names:
            return redirect(url_for("pum_list", pid=pid, error="Tambah minimal satu pengendali UM dulu."))
        wb = openpyxl.load_workbook(io.BytesIO(template["xlsx"]))
        set_pic_columns(wb.active, names)
        fields = pum_layout(wb.active)["fields"]
        if "date" in fields:
            fields["date"].value = datetime.combine(date.today(), datetime.min.time())
        if "purpose" in fields:
            fields["purpose"].value = p["name"]
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        year = date.today().year  # nomor urut per proyek, mulai lagi dari 1 tiap tahun
        seq = con.execute("SELECT COALESCE(MAX(seq), 0) + 1 FROM pum_docs WHERE project_id=? AND year=?",
                          (pid, year)).fetchone()[0]
        doc_id = con.execute("INSERT INTO pum_docs (project_id, xlsx, created_at, updated_at, seq, year)"
                             " VALUES (?,?,?,?,?,?)", (pid, template["xlsx"], now, now, seq, year)).lastrowid
        save_pum(doc_id, wb)
        return redirect(url_for("pum_doc", pid=pid, doc_id=doc_id))
    docs = with_signatures(con.execute("SELECT id, created_at, doc_date, no_pum, periode, total, signatures FROM pum_docs"
                                       " WHERE project_id=? ORDER BY year DESC, seq DESC, id DESC", (pid,)).fetchall())
    unmatched = []
    if p["source"] == "xlsx":
        book, _error = monitoring_book(pid, p)
        unmatched = book["pum_unmatched"] if book else []
    return render_template("pum_list.html", active="monitoring", monitoring_pid=pid, p=p, docs=docs,
                           template=template, unmatched=unmatched, error=request.args.get("error", "")[:200])


def with_signatures(docs):
    """Baris daftar PUM + tanda tangan terpasang: sigs = [{label, name, url}] urut Disetujui, Diajukan."""
    out = []
    for d in docs:
        signed = json.loads(d["signatures"] or "{}")
        sigs = []
        for key, _prefix, label in PUM_SIGN_SLOTS:
            sig = signature_for(signed.get(key, ""))
            sigs.append({"label": label, "name": signed.get(key),
                         "url": url_for("signature_image", sid=sig["id"], v=sig["updated_at"]) if sig else None})
        out.append({**dict(d), "sigs": sigs})
    return out


@app.get("/monitoring/pum")
def pum_all():
    """Sub-menu PUM: semua PUM dari semua proyek monitoring, bisa difilter per proyek."""
    con = db()
    projects = con.execute("SELECT id, name FROM monitoring_projects ORDER BY name").fetchall()
    pid = request.args.get("project", type=int)
    docs = with_signatures(con.execute(
        "SELECT d.id, d.project_id, p.name AS project_name, d.created_at, d.doc_date, d.no_pum, d.periode,"
        " d.total, d.signatures FROM pum_docs d JOIN monitoring_projects p ON p.id = d.project_id"
        + (" WHERE d.project_id = ?" if pid else "")
        + " ORDER BY d.year DESC, p.name, d.seq DESC, d.id DESC", (pid,) if pid else ()).fetchall())
    template = con.execute("SELECT source, imported_at FROM pum_template WHERE id=1").fetchone()
    return render_template("pum_list.html", active="monitoring", monitoring_pid=None, pum_menu=True, p=None,
                           controllers=pum_controllers(),
                           projects=projects, filter_pid=pid, docs=docs, template=template, unmatched=[],
                           error=request.args.get("error", "")[:200])


@app.post("/monitoring/pum/controllers")
@app.post("/monitoring/pum/controllers/<int:cid>")
def pum_controller_save(cid=None):
    """Tambah (tanpa cid) atau ganti nama pengendali UM. Nama wajib, unik, maks 40 huruf."""
    name = request.form.get("name", "").strip()[:40]
    con = db()
    dup = con.execute("SELECT 1 FROM pum_controllers WHERE lower(name)=lower(?) AND id IS NOT ?", (name, cid)).fetchone()
    error = "Nama pengendali wajib diisi." if not name else (f"{name} sudah ada di daftar." if dup else None)
    if not error:
        if cid:
            con.execute("UPDATE pum_controllers SET name=? WHERE id=?", (name, cid))
        else:
            con.execute("INSERT INTO pum_controllers (name) VALUES (?)", (name,))
        con.commit()
    return redirect(url_for("pum_all", **({"error": error} if error else {})) + "#pengendali")


@app.post("/monitoring/pum/controllers/<int:cid>/delete")
def pum_controller_delete(cid):
    db().execute("DELETE FROM pum_controllers WHERE id=?", (cid,))
    db().commit()
    return redirect(url_for("pum_all") + "#pengendali")


@app.get("/monitoring/<int:pid>/pum/<int:doc_id>")
def pum_doc(pid, doc_id):
    p, doc = monitoring_or_404(pid), pum_doc_or_404(pid, doc_id)
    pum = load_pum(doc)
    signed = json.loads(doc["signatures"] or "{}")
    rows, widths = pum_rows(pum, signed)
    slots = [{**x, "signature": signature_for(x["name"]), "signed": signed.get(x["key"], "").lower() == x["name"].lower()}
             for x in pum_sign_slots(pum)]
    periode = pum_layout(pum["wb"].active)["fields"].get("periode")
    range_cell = {"ref": periode.coordinate, "value": periode_range(periode.value)} if periode is not None else None
    return render_template("pum_doc.html", active="monitoring", monitoring_pid=pid, p=p, doc=doc, slots=slots,
                           range_cell=range_cell,
                           sheet={"title": doc["no_pum"] or "PUM baru", "rows": rows, "widths": widths,
                                  "images": 0, "is_list": False},
                           categories=json.loads(doc["categories"]), saved=request.args.get("saved") == "1",
                           error=request.args.get("error", "")[:200])


def pum_rows(pum, signed=None):
    """Baris tabel PUM (read_sheet). Tabel PUM polos: hanya header berwarna (teks putih) yang dipertahankan.
    signed = {slot: nama} dari pum_docs.signatures: gambar tanda tangan tampil di baris kosong tepat di atas nama,
    hanya bila nama di PUM masih sama dengan nama yang menandatangani."""
    ws = pum["wb"].active
    rows, widths = read_sheet(ws, lambda c: book_value(pum, c), ws.max_column, PUM_MAX_ROW)
    for row in rows:
        for c in row["cells"]:
            if c["fill"] and not c["on_fill"]:
                c["fill"] = None
    layout = pum_layout(ws)
    grand, header = layout["grand"], layout["items"][0][0] - 2
    for row in rows:  # baris tanda tangan (di bawah GRAND TOTAL): diberi ruang lebih di tampilan cetak
        row["sign"] = grand is not None and row["n"] > grand.row
    for row in rows:  # isian kepala (label merge A:B, mis. "Tanggal / Date"): tanda ":" sebelum nilainya
        first = row["cells"][0] if row["cells"] else None
        if row["n"] < header and first and first["col"] == 1 and first["colspan"] == 2 and first["text"]:
            first["colon"] = True
    # Gambar tanda tangan mengisi ruang kosong antara label (mis. "Requested by") dan nama: baris-baris itu
    # digabung (rowspan) dan gambar di tengah. Awal ruang sama untuk semua slot supaya tanda tangan sejajar.
    slots = pum_sign_slots(pum)
    if slots:
        by_ref = {c["ref"]: (row, c) for row in rows for c in row["cells"]}
        label_end = max(max((r for r in range(slot["head_row"], slot["name_row"])
                             if ws.cell(r, slot["col"]).value not in (None, "")), default=slot["head_row"]) for slot in slots)
        for slot in slots:
            sig = signature_for(slot["name"]) if (signed or {}).get(slot["key"], "").lower() == slot["name"].lower() else None
            top, bottom = label_end + 1, slot["name_row"] - 1
            hit = by_ref.get(f"{get_column_letter(slot['col'])}{top}")
            if not sig or hit is None or bottom < top:
                continue
            hit[1].update(img=url_for("signature_image", sid=sig["id"], v=sig["updated_at"]), rowspan=bottom - top + 1)
            for r in range(top + 1, bottom + 1):  # sel di bawahnya tertutup oleh sel gabungan
                covered = by_ref.get(f"{get_column_letter(slot['col'])}{r}")
                if covered:
                    covered[0]["cells"].remove(covered[1])
    return rows, widths


# Blok tanda tangan PUM: (slot, awalan label di format, judul)
PUM_SIGN_SLOTS = (("approver", "disetujui oleh", "Disetujui oleh"), ("requester", "diajukan oleh", "Diajukan oleh"))
SIGNATURE_MAX_BYTES = 1024 * 1024


def pum_sign_slots(pum):
    """Blok tanda tangan di bawah GRAND TOTAL: [{key, label, col, name_row, name}]. Nama = sel berisi pertama di bawah
    label pada kolom yang sama (selain 'Requested by'), mis. B57 / E57 (=C8)."""
    ws = pum["wb"].active
    grand = pum_layout(ws)["grand"]
    texts = [(c, str(c.value).strip().lower()) for row in ws.iter_rows(min_row=(grand.row + 1) if grand else 1)
             for c in row if isinstance(c.value, str)]
    slots = []
    for key, prefix, label in PUM_SIGN_SLOTS:
        head = next((c for c, t in texts if t.startswith(prefix)), None)
        if head is None:
            continue
        name_cell = next((ws.cell(r, head.column) for r in range(head.row + 1, ws.max_row + 1)
                          if ws.cell(r, head.column).value not in (None, "")
                          and str(ws.cell(r, head.column).value).strip().lower() not in ("requested by", "approved by")), None)
        if name_cell is None:
            continue
        slots.append({"key": key, "label": label, "col": head.column, "head_row": head.row, "name_row": name_cell.row,
                      "name": str(book_value(pum, name_cell) or "").strip()})
    return slots


def signature_for(name):
    return db().execute("SELECT id, name, mime, updated_at FROM signatures WHERE name=?", (name.strip(),)).fetchone() if name else None


def safe_back(url, fallback):
    """Alamat kembali hanya di dalam aplikasi (cegah redirect ke situs lain)."""
    return url if url and url.startswith("/monitoring/") and "//" not in url else fallback


@app.get("/monitoring/signatures/<int:sid>")
def signature_image(sid):
    row = db().execute("SELECT image, mime FROM signatures WHERE id=?", (sid,)).fetchone()
    if row is None:
        abort(404)
    resp = app.response_class(row["image"], mimetype=row["mime"])
    resp.headers["Cache-Control"] = "private, max-age=86400"  # url memuat versi (?v=updated_at)
    return resp


@app.post("/monitoring/signatures")
def signature_save():
    """Simpan / ganti tanda tangan satu nama: upload PNG/JPG atau gambar di layar (data URL PNG). Maks 1 MB."""
    name = request.form.get("name", "").strip()[:80]
    back = safe_back(request.form.get("back"), url_for("pum_all"))
    data, mime = None, None
    upload = request.files.get("image")
    if upload and upload.filename:
        data = upload.read(SIGNATURE_MAX_BYTES + 1)
    elif request.form.get("drawn", "").startswith("data:image/png;base64,"):
        try:
            data = base64.b64decode(request.form["drawn"].split(",", 1)[1], validate=True)
        except ValueError:
            data = None
    if data and data.startswith(b"\x89PNG\r\n\x1a\n"):
        mime = "image/png"
    elif data and data.startswith(b"\xff\xd8\xff"):
        mime = "image/jpeg"
    sep = "&" if "?" in back else "?"
    if not name:
        return redirect(back + sep + "error=Nama+penanda+tangan+kosong.")
    if not data or not mime:
        return redirect(back + sep + "error=Pilih+gambar+PNG/JPG+atau+gambar+tanda+tangan+dulu.")
    if len(data) > SIGNATURE_MAX_BYTES:
        return redirect(back + sep + "error=Gambar+tanda+tangan+maksimal+1+MB.")
    db().execute("INSERT INTO signatures (name, image, mime, updated_at) VALUES (?,?,?,?)"
                 " ON CONFLICT(name) DO UPDATE SET image=excluded.image, mime=excluded.mime, updated_at=excluded.updated_at",
                 (name, data, mime, datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")))
    db().commit()
    return redirect(back + sep + "saved=1")


@app.post("/monitoring/signatures/<int:sid>/delete")
def signature_delete(sid):
    db().execute("DELETE FROM signatures WHERE id=?", (sid,))
    db().commit()
    return redirect(safe_back(request.form.get("back"), url_for("pum_all")))


@app.post("/monitoring/<int:pid>/pum/<int:doc_id>/sign")
def pum_sign(pid, doc_id):
    """Tandatangani (action=sign) atau lepas tanda tangan (unsign) satu slot PUM dengan tanda tangan milik nama di slot."""
    doc = pum_doc_or_404(pid, doc_id)
    slot = next((x for x in pum_sign_slots(load_pum(doc)) if x["key"] == request.form.get("slot")), None)
    if slot is None:
        abort(400)
    signed = json.loads(doc["signatures"] or "{}")
    back = url_for("pum_doc", pid=pid, doc_id=doc_id)
    if request.form.get("action") == "sign":
        if not signature_for(slot["name"]):
            return redirect(back + "?error=Belum+ada+tanda+tangan+untuk+" + slot["name"].replace(" ", "+") + ".")
        signed[slot["key"]] = slot["name"]
    else:
        signed.pop(slot["key"], None)
    db().execute("UPDATE pum_docs SET signatures=? WHERE id=?", (json.dumps(signed), doc_id))
    db().commit()
    return redirect(back + "?saved=1#ttd")


@app.get("/monitoring/<int:pid>/pum/<int:doc_id>/pdf")
def pum_pdf(pid, doc_id):
    """Halaman cetak A4 satu PUM; dialog cetak terbuka otomatis (Simpan sebagai PDF). Judul = nama file PDF."""
    monitoring_or_404(pid)
    doc = pum_doc_or_404(pid, doc_id)
    rows, widths = pum_rows(load_pum(doc), json.loads(doc["signatures"] or "{}"))
    return render_template("pum_print.html", doc=doc, rows=rows, widths=widths)


@app.post("/monitoring/<int:pid>/pum/<int:doc_id>/cell")
def pum_cell(pid, doc_id):
    """Edit satu sel input PUM (A1:G58); sel rumus terkunci."""
    wb = openpyxl.load_workbook(io.BytesIO(pum_doc_or_404(pid, doc_id)["xlsx"]))
    ws = wb.active
    ref = request.form.get("cell", "").upper()
    back = lambda **kw: redirect(url_for("pum_doc", pid=pid, doc_id=doc_id, **kw) + f"#c-{ref}")
    m = re.fullmatch(r"([A-Z]{1,3})([0-9]{1,6})", ref)
    if not m or column_index_from_string(m.group(1)) > ws.max_column or int(m.group(2)) > PUM_MAX_ROW:
        return back(error="Sel di luar area PUM.")
    cell = ws[ref]
    no_pum = pum_layout(ws)["fields"].get("no_pum")
    if no_pum is not None and cell.coordinate == no_pum.coordinate:
        return back(error="No. PUM dibuat otomatis dari nomor urut dan Tanggal PUM.")
    if is_formula(cell):
        return back(error=f"Sel {ref} berisi rumus dan dihitung otomatis, tidak bisa diedit.")
    periode = pum_layout(ws)["fields"].get("periode")
    if periode is not None and cell.coordinate == periode.coordinate and "range_from" in request.form:
        try:  # Periode diisi lewat 2 kalender -> teks "11 Nov - 17 Nov 2026"
            start, end = (date.fromisoformat(request.form.get(k, "")) for k in ("range_from", "range_to"))
        except ValueError:
            return back(error="Isi tanggal mulai dan selesai periode.")
        if end < start:
            return back(error="Tanggal selesai periode tidak boleh sebelum tanggal mulai.")
        new = periode_text(start, end)
    else:
        new = parse_cell_input(request.form.get("value", ""), cell)
    cell.value = new
    if isinstance(new, str):
        cell.data_type = "s"  # teks berawalan "=" tetap teks
    save_pum(doc_id, wb)
    return back(saved=1)


def periode_text(start, end):
    """Rentang tanggal -> teks Periode PUM: '11 Nov - 17 Nov 2026', beda tahun: '28 Des 2026 - 3 Jan 2027'."""
    left = f"{start.day} {MONTHS[start.month - 1]}" + (f" {start.year}" if start.year != end.year else "")
    return f"{left} - {end.day} {MONTHS[end.month - 1]} {end.year}"


def periode_range(text):
    """Kebalikan periode_text untuk mengisi kalender: (mulai, selesai) ISO, atau ('', '') bila formatnya lain."""
    m = re.fullmatch(r"\s*(\d{1,2}) (\w{3})(?: (\d{4}))? - (\d{1,2}) (\w{3}) (\d{4})\s*", str(text or ""))
    month = {name.lower(): i + 1 for i, name in enumerate(MONTHS)}
    try:
        d1, m1, y1, d2, m2, y2 = m.groups()
        end = date(int(y2), month[m2.lower()], int(d2))
        return date(int(y1 or y2), month[m1.lower()], int(d1)).isoformat(), end.isoformat()
    except (AttributeError, KeyError, ValueError):
        return "", ""


@app.post("/monitoring/<int:pid>/pum/<int:doc_id>/delete")
def pum_delete(pid, doc_id):
    pum_doc_or_404(pid, doc_id)
    db().execute("DELETE FROM pum_docs WHERE id=?", (doc_id,))
    db().commit()
    _pums.pop(doc_id, None)
    return redirect(url_for("pum_list", pid=pid))

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5010,debug=True)
