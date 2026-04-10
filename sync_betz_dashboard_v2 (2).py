from __future__ import annotations

import csv
import html
import time
from datetime import datetime
from pathlib import Path

SOURCE_DIR = Path(r"\\10.1.1.225\spooltext\Productivity")
CSV_FILENAME = "COLL_STATS.csv"
CSV_PATH = SOURCE_DIR / CSV_FILENAME

OUTPUT_HTML = SOURCE_DIR / "dashboard_betz_ar_live.html"
LOCAL_OUTPUT_HTML = Path(r"C:\Users\Public\dashboard_betz_ar_live.html")
CHECK_EVERY_SECONDS = 5
AUTO_REFRESH_SECONDS = 900

# ---------------------------------------------------------------------------
# Mapeo de FAC -> departamento  (extraído del Excel Productivity Report BETZ&AR 2026)
# ---------------------------------------------------------------------------
FAC_TO_INFO = {
    # PT CONTACT
    "515": {"dept": "pt", "name": "Auri"},
    "312": {"dept": "pt", "name": "Joanna"},
    "309": {"dept": "pt", "name": "Joel R"},
    "508": {"dept": "pt", "name": "Daniel"},
    "518": {"dept": "pt", "name": "Carlos"},
    "311": {"dept": "pt", "name": "Bianca"},
    "124": {"dept": "pt", "name": "Marcos"},
    "519": {"dept": "pt", "name": "Domenica"},
    "127": {"dept": "pt", "name": "Pending"},
    "129": {"dept": "pt", "name": "Pending"},
    "122": {"dept": "pt", "name": "Pending"},

    # HOSPITAL FOLLOW UP
    "310": {"dept": "hospital", "name": "Genesis"},
    "125": {"dept": "hospital", "name": "Dyanna"},
    "113": {"dept": "hospital", "name": "Karla"},
    "126": {"dept": "hospital", "name": "Ivett"},
    "118": {"dept": "hospital", "name": "Juan Jose R."},

    # DENIALS
    "533": {"dept": "denials", "name": "Galo Pacheco"},
    "506": {"dept": "denials", "name": "Noemi"},

    # MEDICAID
    "130": {"dept": "medicaid", "name": "Emilio Cali"},
    "103": {"dept": "medicaid", "name": "Ronald"},
    "308": {"dept": "medicaid", "name": "Allison"},

    # NO RESPONSE
    "120": {"dept": "noresponse", "name": "Joseline"},
    "123": {"dept": "noresponse", "name": "Daniel H"},
    "313": {"dept": "noresponse", "name": "Luis"},
    "507": {"dept": "noresponse", "name": "Pending"},
    "511": {"dept": "noresponse", "name": "Pending"},
    "516": {"dept": "noresponse", "name": "Pending"},
    "536": {"dept": "noresponse", "name": "Pending"},

    # NYP UNDERPAYMENTS
    "401": {"dept": "nypunder", "name": "Natalia"},
    "402": {"dept": "nypunder", "name": "Paulina"},
    "404": {"dept": "nypunder", "name": "Naomi"},
    "405": {"dept": "nypunder", "name": "Jairon"},
    "406": {"dept": "nypunder", "name": "Valeria"},
    "408": {"dept": "nypunder", "name": "Veronica"},
    "410": {"dept": "nypunder", "name": "Angy"},
    "411": {"dept": "nypunder", "name": "Raiza"},
    "412": {"dept": "nypunder", "name": "Pending"},
    "413": {"dept": "nypunder", "name": "Pending"},
    "415": {"dept": "nypunder", "name": "Pending"},

    # JAMAICA / FLUSHING PB FU
    "517": {"dept": "jamaica", "name": "Jorge"},

    # NF / COMP
    "307": {"dept": "nfcomp", "name": "Ana Jaramillo"},
    "951": {"dept": "nfcomp", "name": "Pending"},
    "924": {"dept": "nfcomp", "name": "Pending"},

    # ROSWELL
    "314": {"dept": "roswell", "name": "Sergio"},
    "302": {"dept": "roswell", "name": "Milena"},
    "563": {"dept": "roswell", "name": "Edgar"},
    "532": {"dept": "roswell", "name": "Diego"},
    "534": {"dept": "roswell", "name": "Pending"},
    "303": {"dept": "roswell", "name": "Pending"},
    "535": {"dept": "roswell", "name": "Pending"},
    "306": {"dept": "roswell", "name": "Pending"},
    "305": {"dept": "roswell", "name": "Pending"},
    "301": {"dept": "roswell", "name": "Pending"},

    # PB CASH POSTING
    "509": {"dept": "cashposting", "name": "Zahid"},
    "531": {"dept": "cashposting", "name": "Andres"},

    # BILLING
    "523": {"dept": "billing", "name": "Samuel Yepez"},
    "530": {"dept": "billing", "name": "Maria Paula"},
    "109": {"dept": "billing", "name": "Stephanie"},
    "503": {"dept": "billing", "name": "Kerlly"},
    "117": {"dept": "billing", "name": "Franz"},
    "116": {"dept": "billing", "name": "Ivan"},
    "315": {"dept": "billing", "name": "Dana"},
    "304": {"dept": "billing", "name": "Yolanny"},
    "114": {"dept": "billing", "name": "Gustavo"},
}

DEPT_META = {
    "pt":         {"label": "PT Contact",              "icon": "📞", "color": "#2563eb"},
    "hospital":   {"label": "Hospital Follow Up",      "icon": "🏥", "color": "#d97706"},
    "denials":    {"label": "Denials WQ",              "icon": "🚫", "color": "#7c3aed"},
    "medicaid":   {"label": "Medicaid",                "icon": "💊", "color": "#e11d48"},
    "noresponse": {"label": "No Response WQ",          "icon": "🔕", "color": "#0f766e"},
    "nypunder":   {"label": "NYP Underpayments",       "icon": "📉", "color": "#16a34a"},
    "jamaica":    {"label": "Jamaica / Flushing PB FU","icon": "📋", "color": "#0891b2"},
    "nfcomp":     {"label": "NF / Comp Underpayments", "icon": "🧷", "color": "#65a30d"},
    "roswell":    {"label": "Roswell",                 "icon": "🧬", "color": "#64748b"},
    "cashposting":{"label": "PB Cash Posting",         "icon": "💵", "color": "#ea580c"},
    "billing":    {"label": "Billing",                 "icon": "🧾", "color": "#4f46e5"},
}

DEPT_ORDER = ["pt","hospital","denials","medicaid","noresponse",
              "nypunder","jamaica","nfcomp","roswell","cashposting","billing"]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def as_int(v) -> int:
    try:
        return int(v) if v not in (None, "", "0") else 0
    except Exception:
        return 0


def read_csv(csv_path: Path):
    """
    Lee COLL_STATS.csv y devuelve lista de dicts por fila.
    Columnas: A=fac, B=mes, C=dia, D=anio, E=calls, F=contacts,
              G=accounts, H=letters, I=time
    """
    rows = []
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        for line in reader:
            if not line:
                continue
            # Quitar comillas residuales si las hay
            line = [c.strip().strip('"') for c in line]
            if len(line) < 9:
                continue
            fac = line[0].strip()
            if not fac:
                continue
            rows.append({
                "fac":      fac,
                "mes":      line[1],
                "dia":      line[2],
                "anio":     line[3],
                "calls":    as_int(line[4]),
                "contacts": as_int(line[5]),
                "accounts": as_int(line[6]),
                "letters":  as_int(line[7]),
                "time":     as_int(line[8]),
            })
    return rows


def build_departments(rows):
    dept_data: dict[str, list[dict]] = {k: [] for k in DEPT_META}

    for row in rows:
        fac = row["fac"]
        info = FAC_TO_INFO.get(fac)
        if info is None:
            continue

        dept_key = info["dept"]
        dept_data[dept_key].append({
            "name": info["name"],
            "facs": fac,
            "accounts": row["accounts"],
            "calls": row["calls"],
            "contacts": row["contacts"],
            "time": row["time"],
            "letters": row["letters"],
            "pct": 0.0,
        })

    departments = []
    for key in DEPT_ORDER:
        reps = dept_data.get(key, [])
        reps = [r for r in reps if any([r["accounts"], r["calls"], r["contacts"], r["time"], r["letters"]])]
        meta = DEPT_META[key]
        totals = {
            "accounts": sum(r["accounts"] for r in reps),
            "calls": sum(r["calls"] for r in reps),
            "contacts": sum(r["contacts"] for r in reps),
            "time": sum(r["time"] for r in reps),
            "letters": sum(r["letters"] for r in reps),
            "pct": 0.0,
        }
        departments.append({
            "key": key,
            "label": meta["label"],
            "icon": meta["icon"],
            "color": meta["color"],
            "reps": reps,
            "totals": totals,
        })
    return departments

    departments = []
    for key in DEPT_ORDER:
        reps = dept_data.get(key, [])
        # Filtrar filas con todo en cero
        reps = [r for r in reps if any([r["accounts"], r["calls"], r["contacts"], r["time"], r["letters"]])]
        meta = DEPT_META[key]
        totals = {
            "accounts": sum(r["accounts"] for r in reps),
            "calls":    sum(r["calls"]    for r in reps),
            "contacts": sum(r["contacts"] for r in reps),
            "time":     sum(r["time"]     for r in reps),
            "letters":  sum(r["letters"]  for r in reps),
            "pct":      0.0,
        }
        departments.append({
            "key":    key,
            "label":  meta["label"],
            "icon":   meta["icon"],
            "color":  meta["color"],
            "reps":   reps,
            "totals": totals,
        })
    return departments


def get_report_date(rows) -> str:
    """Usa la fecha del primer registro del CSV."""
    for row in rows:
        try:
            return f"{row['mes']}/{row['dia']}/{row['anio']}"
        except Exception:
            pass
    return datetime.now().strftime("%m/%d/%Y")


# ---------------------------------------------------------------------------
# HTML rendering (igual que antes, adaptado)
# ---------------------------------------------------------------------------

def rank_cards_html(departments):
    order = sorted(departments, key=lambda d: (-d["totals"]["accounts"], d["label"].lower()))
    out = []
    max_accounts = max([d["totals"]["accounts"] for d in departments] + [1])
    for d in order:
        pct = round(d["totals"]["accounts"] / max_accounts * 100, 1)
        active_reps = len([r for r in d["reps"] if r["accounts"] > 0 or r["calls"] > 0])
        out.append(f"""
        <div class="dept-card" onclick="goSection('{html.escape(d['key'])}')">
          <div class="dept-top">
            <div class="dept-icon" style="background:{d['color']}15;color:{d['color']}">{d['icon']}</div>
            <div class="dept-accounts">{d['totals']['accounts']}</div>
          </div>
          <div class="dept-name">{html.escape(d['label'])}</div>
          <div class="dept-sub">{active_reps} Active FACS · Calls {d['totals']['calls']} · Contacts {d['totals']['contacts']}</div>
          <div class="dept-bar"><div class="dept-fill" style="width:{pct}%;background:{d['color']}"></div></div>
        </div>
        """)
    return "".join(out)


def sidebar_html(departments):
    return "".join(
        f"""<button class="nav-item" id="nav-{html.escape(d['key'])}" onclick="goSection('{html.escape(d['key'])}')"><span class="nav-dot" style="background:{d['color']}"></span>{html.escape(d['label'])}</button>"""
        for d in departments
    )


def reps_rows_html(reps):
    rows = []
    for r in reps:
        rows.append(
            f"<tr>"
            f"<td>{html.escape(r['name'])}</td>"
            f"<td>{html.escape(r['facs'])}</td>"
            f"<td>{r['accounts']}</td>"
            f"<td>{r['calls']}</td>"
            f"<td>{r['contacts']}</td>"
            f"<td>{r['letters']}</td>"
            f"<td>{r['time']}</td>"
            f"</tr>"
        )
    return "".join(rows)


def section_html(d):
    total = d["totals"]
    active_reps = len([r for r in d["reps"] if r["accounts"] > 0 or r["calls"] > 0])
    return f"""
    <section id="sec-{html.escape(d['key'])}" class="section">
      <div class="page-header">
        <h2>{d['icon']} {html.escape(d['label'])}</h2>
        <p>{active_reps} FACS activos · Accounts {total['accounts']} · Calls {total['calls']} · Contacts {total['contacts']}</p>
      </div>
      <div class="kpi-grid">
        <div class="kpi-card"><div class="kpi-label">Accounts</div><div class="kpi-value">{total['accounts']}</div></div>
        <div class="kpi-card"><div class="kpi-label">Calls</div><div class="kpi-value">{total['calls']}</div></div>
        <div class="kpi-card"><div class="kpi-label">Contacts</div><div class="kpi-value">{total['contacts']}</div></div>
        <div class="kpi-card"><div class="kpi-label">Time (min)</div><div class="kpi-value">{total['time']}</div></div>
      </div>
      <div class="card">
        <div class="card-title">Employee / FACS Detail</div>
        <table>
          <thead>
            <tr><th>Name</th><th>FACS</th><th>Accounts</th><th>Calls</th><th>Contacts</th><th>Letters</th><th>Time</th></tr>
          </thead>
          <tbody>{reps_rows_html(d['reps'])}</tbody>
          <tfoot>
            <tr>
              <td>Total</td>
              <td>—</td>
              <td>{total['accounts']}</td>
              <td>{total['calls']}</td>
              <td>{total['contacts']}</td>
              <td>{total['letters']}</td>
              <td>{total['time']}</td>
            </tr>
          </tfoot>
        </table>
      </div>
    </section>
    """


def render_html(departments, report_date, source_name):
    total_accounts = sum(d["totals"]["accounts"] for d in departments)
    total_calls    = sum(d["totals"]["calls"]    for d in departments)
    total_contacts = sum(d["totals"]["contacts"] for d in departments)
    total_time     = sum(d["totals"]["time"]     for d in departments)
    top_dept = max(departments, key=lambda d: d["totals"]["accounts"]) if departments else {"label": "—", "totals": {"accounts": 0}}
    sections = "".join(section_html(d) for d in departments)
    max_accts = max([d["totals"]["accounts"] for d in departments] + [1])

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="refresh" content="{AUTO_REFRESH_SECONDS}">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>BETZ AR Dashboard Live</title>
<style>
*{{box-sizing:border-box}} body{{margin:0;font-family:Inter,Segoe UI,Arial,sans-serif;background:#f5f7fb;color:#172033}}
.layout{{display:flex;min-height:100vh}}
.sidebar{{width:270px;background:#0f172a;color:#e5e7eb;padding:22px 16px;position:fixed;left:0;top:0;bottom:0;overflow:auto}}
.brand{{font-weight:800;font-size:20px;letter-spacing:.02em}}
.brand-sub{{font-size:12px;color:#94a3b8;margin-top:6px}}
.nav-label{{font-size:11px;text-transform:uppercase;letter-spacing:.12em;color:#94a3b8;margin:26px 10px 10px}}
.nav-item{{width:100%;border:none;background:transparent;color:#e5e7eb;display:flex;align-items:center;gap:10px;padding:11px 12px;border-radius:12px;cursor:pointer;text-align:left;font-weight:600}}
.nav-item:hover,.nav-item.active{{background:#1e293b}}
.nav-dot{{width:10px;height:10px;border-radius:999px;display:inline-block;flex:0 0 auto}}
.main{{margin-left:270px;flex:1;padding:28px}}
.hero{{background:linear-gradient(135deg,#0f172a,#1d4ed8);border-radius:22px;padding:26px 28px;color:white;box-shadow:0 20px 50px rgba(37,99,235,.18)}}
.hero h1{{margin:0;font-size:28px}} .hero p{{margin:8px 0 0;color:#dbeafe}}
.hero-meta{{margin-top:14px;font-size:13px;color:#bfdbfe}}
.kpi-grid{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;margin:22px 0}}
.kpi-card,.card,.dept-card{{background:white;border:1px solid #e2e8f0;border-radius:18px;box-shadow:0 8px 24px rgba(15,23,42,.05)}}
.kpi-card{{padding:20px}} .kpi-label{{font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:#64748b;font-weight:700}}
.kpi-value{{margin-top:8px;font-size:32px;font-weight:800}}
.grid-main{{display:grid;grid-template-columns:1.45fr 1fr;gap:18px;align-items:start}}
.card{{padding:20px}}
.card-title{{font-size:15px;font-weight:800;margin-bottom:14px}}
.dept-grid{{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px}}
.dept-card{{padding:16px;cursor:pointer;transition:transform .15s ease, box-shadow .15s ease}}
.dept-card:hover{{transform:translateY(-2px);box-shadow:0 12px 28px rgba(15,23,42,.1)}}
.dept-top{{display:flex;justify-content:space-between;align-items:center}}
.dept-icon{{width:40px;height:40px;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:18px}}
.dept-accounts{{font-size:28px;font-weight:800}}
.dept-name{{font-weight:800;margin-top:12px}}
.dept-sub{{font-size:12px;color:#64748b;margin-top:4px;min-height:32px}}
.dept-bar{{height:7px;background:#eef2f7;border-radius:999px;margin-top:12px;overflow:hidden}}
.dept-fill{{height:100%;border-radius:999px}}
.summary-list{{display:flex;flex-direction:column;gap:12px}}
.summary-row{{display:grid;grid-template-columns:1fr auto auto;gap:12px;align-items:center;font-size:14px}}
.summary-row strong{{font-size:13px}}
.summary-track{{width:140px;height:8px;background:#eef2f7;border-radius:999px;overflow:hidden}}
.summary-fill{{height:100%;border-radius:999px}}
.section{{display:none;margin-top:28px}} .section.active{{display:block}}
.page-header h2{{margin:0;font-size:24px}} .page-header p{{margin:6px 0 18px;color:#64748b}}
table{{width:100%;border-collapse:collapse}} th,td{{padding:10px 8px;border-bottom:1px solid #e2e8f0;text-align:right;font-size:13px}}
th:first-child,td:first-child{{text-align:left}}
thead th{{font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:#64748b}} tfoot td{{font-weight:800}}
@media (max-width:1200px){{.dept-grid{{grid-template-columns:repeat(2,minmax(0,1fr))}} .grid-main{{grid-template-columns:1fr}}}}
@media (max-width:900px){{.sidebar{{position:static;width:100%}} .main{{margin-left:0}} .layout{{display:block}} .kpi-grid{{grid-template-columns:repeat(2,minmax(0,1fr))}} .dept-grid{{grid-template-columns:1fr}}}}
</style>
</head>
<body>
<div class="layout">
  <aside class="sidebar">
    <div class="brand">BETZ AR</div>
    <div class="brand-sub">Live Productivity Dashboard</div>
    <div class="nav-label">Overview</div>
    <button class="nav-item active" id="nav-overview" onclick="goSection('overview')"><span class="nav-dot" style="background:#38bdf8"></span>Dashboard Overview</button>
    <div class="nav-label">Departments</div>
    {sidebar_html(departments)}
  </aside>
  <main class="main">
    <div class="hero">
      <h1>Accounts-first productivity view</h1>
      <p>Synced with COLL_STATS.csv · Grouped by FACS number</p>
      <div class="hero-meta">Report date: {html.escape(report_date)} · Source: {html.escape(source_name)} · Auto refresh: {AUTO_REFRESH_SECONDS}s</div>
    </div>

    <section id="sec-overview" class="section active">
      <div class="kpi-grid">
        <div class="kpi-card"><div class="kpi-label">Departments</div><div class="kpi-value">{len(departments)}</div></div>
        <div class="kpi-card"><div class="kpi-label">Total Accounts</div><div class="kpi-value">{total_accounts}</div></div>
        <div class="kpi-card"><div class="kpi-label">Total Calls</div><div class="kpi-value">{total_calls}</div></div>
        <div class="kpi-card"><div class="kpi-label">Top Department</div><div class="kpi-value" style="font-size:22px">{html.escape(top_dept['label'])}</div><div style="margin-top:6px;color:#64748b;font-size:13px">{top_dept['totals']['accounts']} accounts</div></div>
      </div>

      <div class="card">
        <div class="card-title">Accounts by Department</div>
        <div class="dept-grid">
          {rank_cards_html(departments)}
        </div>
      </div>

      <div class="grid-main">
        <div class="card">
          <div class="card-title">Executive Totals</div>
          <div class="kpi-grid" style="margin:0">
            <div class="kpi-card"><div class="kpi-label">Accounts</div><div class="kpi-value">{total_accounts}</div></div>
            <div class="kpi-card"><div class="kpi-label">Calls</div><div class="kpi-value">{total_calls}</div></div>
            <div class="kpi-card"><div class="kpi-label">Contacts</div><div class="kpi-value">{total_contacts}</div></div>
            <div class="kpi-card"><div class="kpi-label">Time (min)</div><div class="kpi-value">{total_time}</div></div>
          </div>
        </div>
        <div class="card">
          <div class="card-title">Department Ranking</div>
          <div class="summary-list">
            {''.join(
              f"<div class='summary-row'><strong>{html.escape(d['label'])}</strong>"
              f"<div class='summary-track'><div class='summary-fill' style='width:{round(d['totals']['accounts']/max_accts*100)}%;background:{d['color']}'></div></div>"
              f"<span>{d['totals']['accounts']}</span></div>"
              for d in sorted(departments, key=lambda x: (-x['totals']['accounts'], x['label']))
            )}
          </div>
        </div>
      </div>
    </section>

    {sections}
  </main>
</div>
<script>
function goSection(key) {{
  document.querySelectorAll('.section').forEach(x => x.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(x => x.classList.remove('active'));
  const sec = document.getElementById('sec-' + key);
  const nav = document.getElementById('nav-' + key);
  if (sec) sec.classList.add('active');
  if (nav) nav.classList.add('active');
}}
</script>
</body>
</html>"""


# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------

def build_once():
    rows = read_csv(CSV_PATH)
    report_date = get_report_date(rows)
    departments = build_departments(rows)
    html_text = render_html(departments, report_date, CSV_FILENAME)
    OUTPUT_HTML.write_text(html_text, encoding="utf-8")
    try:
        LOCAL_OUTPUT_HTML.write_text(html_text, encoding="utf-8")
    except Exception:
        pass
    return departments


def main():
    last_run_slot = None
    print("BETZ dashboard sync iniciado (modo CSV)...")

    while True:
        try:
            now = datetime.now()
            minute = now.minute
            slot = (now.year, now.month, now.day, now.hour, minute)

            # Solo correr en :01, :16, :31, :46
            if minute in {1, 16, 31, 46} and slot != last_run_slot:
                departments = build_once()
                last_run_slot = slot
                print(
                    f"[OK] Dashboard actualizado desde: {CSV_FILENAME} | "
                    f"departamentos activos: {sum(1 for d in departments if d['totals']['accounts'] > 0)} | "
                    f"{now.strftime('%Y-%m-%d %H:%M:%S')}"
                )

        except Exception as e:
            print(f"[ERROR] {e}")

        time.sleep(CHECK_EVERY_SECONDS)