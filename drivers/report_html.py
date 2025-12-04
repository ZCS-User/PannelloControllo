# drivers/report_html.py
import os, glob, tempfile, shutil, subprocess, pathlib
import pandas as pd
import plotly.express as px
from jinja2 import Environment, Template
import numpy as np


def _pick_col(df, candidates):
    for name in candidates:
        if name in df.columns:
            return name
    return None

def _sum_dc_power(df, allowed):
    total = None
    for ch in allowed:
        name = f"Power {ch} [W]"
        if name in df.columns:
            s = pd.to_numeric(df[name], errors="coerce")
            total = s if total is None else (total.add(s, fill_value=0))
    return total


def _headless_pdf(html_path: str, pdf_path: str):
    """Genera PDF con Edge/Chrome headless."""
    candidates = [
        shutil.which("msedge"), shutil.which("msedge.exe"),
        shutil.which("chrome"), shutil.which("chrome.exe"),
        shutil.which("chromium"), shutil.which("chromium-browser"),
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    bin_path = next((c for c in candidates if c and os.path.exists(c)), None)
    if not bin_path:
        return False, "Browser headless non trovato (Edge/Chrome)."

    uri = pathlib.Path(html_path).resolve().as_uri()
    out_abs = os.path.abspath(pdf_path)

    cmds = [
        [bin_path, "--headless=new", "--disable-gpu", f"--print-to-pdf={out_abs}", uri],
        [bin_path, "--headless", "--disable-gpu", f"--print-to-pdf={out_abs}", uri],
    ]
    last_err = ""
    for cmd in cmds:
        try:
            res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=90)
            if res.returncode == 0 and os.path.isfile(out_abs):
                return True, ""
            last_err = (res.stderr.decode(errors="ignore") or res.stdout.decode(errors="ignore"))[:500]
        except Exception as e:
            last_err = str(e)
    return False, last_err


def _headless_pdf(html_path: str, pdf_path: str) -> tuple[bool, str]:
    """Prova a generare PDF usando Edge/Chrome headless."""
    candidates = [
        shutil.which("msedge"), shutil.which("msedge.exe"),
        shutil.which("chrome"), shutil.which("chrome.exe"),
        shutil.which("chromium"), shutil.which("chromium-browser"),
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    bin_path = next((c for c in candidates if c and os.path.exists(c)), None)
    if not bin_path:
        return False, "Browser headless non trovato (Edge/Chrome)."

    uri = pathlib.Path(html_path).resolve().as_uri()
    out_abs = os.path.abspath(pdf_path)
    cmd = [bin_path, "--headless=new", "--disable-gpu", "--print-to-pdf="+out_abs, uri]
    try:
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=90)
        if res.returncode == 0 and os.path.isfile(out_abs):
            return True, ""
        return False, (res.stderr.decode(errors="ignore") or res.stdout.decode(errors="ignore"))[:400]
    except Exception as e:
        return False, str(e)


def _allowed_chans_from_template(template_path: str):
    base = os.path.basename(template_path).upper()
    allowed = set()
    for tag in ("DC1","DC2","DC3"):
        if tag in base:
            allowed.add(tag)
    # se non c'è nessun tag, consenti tutti (compatibilità vecchi template)
    return allowed or {"DC1","DC2","DC3"}


def _basename_filter(p):
    try:
        return os.path.basename(p)
    except Exception:
        return p


# --- Regole di valutazione per i test ---
TEST_SPECS = {
    # match per nome file (case-insensitive) → regole
    "curva mppt": {
        "mode": "curva_mppt",
        "ref_db_key": "PCH|MAX POUT|MAX POUT [kW]",                        # colonna nel DB inverter
        "tol_percent": 5.0,                         # ± %
        "meas_column": "Active Output Power [W]",  # colonna nel log
        "meas_reduce": "max",                       # max/mean/last
        "title": "Report Test – Curva MPPT",
        "ref_unit": "W",       # PCH nel DB è in kW
        "meas_unit": "kW",       # il log è in W
    },
    "ciclo batteria": {
        "mode": "battery_cycle",
        "tol_percent": 5.0,
        "pbat_col": "Charge/Discharge Power [kW]",
        "pbat_unit": "kW",       # se nel DB è in kW, cambia in "kW"
        "ac_col": "Active Output Power [W]",
        "settle_min_s": 5.0,
        "settle_ratio": 0.2,
        "title": "Report Test – Ciclo Batteria",
    },
}


# --- Unit helpers ------------------------------------------------------------
def _unit_scale(src: str|None, dst: str|None) -> float:
    s = (src or "").strip().lower()
    d = (dst or "").strip().lower()
    if not s or not d or s == d:
        return 1.0
    table = {
        ("kw","w"): 1000.0,
        ("w","kw"): 0.001,
    }
    return table.get((s,d), 1.0)

def _fmt_value(v, unit: str|None):
    if v is None:
        return "n/d"
    try:
        v = float(v)
    except Exception:
        return "n/d"
    suf = f" {unit}" if unit else ""
    return f"{v:,.3f}{suf}".replace(",", "X").replace(".", ",").replace("X",".")
# ----------------------------------------------------------------------------


def _find_batt_power_col(df):
    cand = ["Charge/Discharge Power [W]", "Battery Power [W]", "PBAT [W]"]
    for c in cand:
        if c in df.columns: return c
    # fallback: cerca qualcosa che contenga "batt" e "power"
    for c in df.columns:
        lc = c.lower()
        if "batt" in lc and "power" in lc: return c
    return None


def _build_time_axis(df):
    """Ritorna (t_s, has_ts): array di secondi da inizio e flag se c'era timestamp.
       Se non c'è timestamp, usa dt uniforme dalla mediana delle differenze o 1 s."""
    import pandas as pd
    import numpy as np
    # prova colonne tipiche
    for col in ["timestamp","time","Time","Timestamp"]:
        if col in df.columns:
            ts = pd.to_datetime(df[col], errors="coerce")
            if ts.notna().any():
                t0 = ts.dropna().iloc[0]
                t_s = (ts - t0).dt.total_seconds().ffill().fillna(0.0).to_numpy()
                return t_s, True
    # fallback: indice con passo stimato
    n = len(df)
    if n <= 1: return (np.zeros(n), False)
    # stima dt da una eventuale colonna 'ms'/'s' o dal numero di righe
    dt = 1.0
    try:
        dt = float(np.median(np.diff(df.index.to_numpy()))) if n>2 else 1.0
        if not np.isfinite(dt) or dt<=0: dt = 1.0
    except Exception:
        dt = 1.0
    import numpy as np
    return (np.arange(n, dtype=float)*dt, False)


def _read_template_steps(template_path, sn_for_db):
    """Legge il template e ritorna lista di step: dict con Psp(W), sign, durata(s)."""
    import pandas as pd
    df_t = pd.read_excel(template_path)
    # DB fallback per "P BAT" nel template
    pbat_db = _db_get_value_for(sn_for_db, "P BAT")
    steps = []
    for i in range(len(df_t)):
        raw = df_t.loc[i, "potenza batteria"]
        try:
            if str(raw).strip().upper() == "P BAT":
                psp = float(pbat_db or 0.0)
            else:
                psp = float(str(raw).replace(",", "."))
        except Exception:
            psp = float(pbat_db or 0.0)
        mode = str(df_t.loc[i, "scarica/carica"]).strip().lower()
        sign = -1.0 if mode == "scarica" else 1.0
        try:
            dur = float(str(df_t.loc[i, "tempo"]).replace(",", "."))
        except Exception:
            dur = 1.0
        steps.append({"psp": sign*psp, "dur": max(0.1, dur)})
    return steps


def _norm(s: str) -> str:
    return (s or "").strip().lower()


def _guess_spec_from_template(template_path: str):
    base = os.path.basename(template_path)
    nb = _norm(os.path.splitext(base)[0])
    for key, spec in TEST_SPECS.items():
        if key in nb:
            return spec
    # default “generico”
    return {
        "ref_db_key": None,
        "tol_percent": None,
        "meas_column": None,
        "meas_reduce": "max",
        "title": f"Report Test – {os.path.splitext(base)[0]}",
    }


# --- Parse SN semplificato per DB lookup ---
def _parse_sn_family_model(sn: str):
    sn = (sn or "").strip()
    if len(sn) == 14:
        family = sn[:3]     # prime 3 della definizione (5)
        model  = sn[5:8]    # 6–8 (0-indexed 5..8)
    elif len(sn) == 20:
        family = sn[:6]     # **correzione**: primi 5 per famiglia nel 20-cifre
        model  = sn[6:9]    # 7–9 (0-indexed 6..9)
    else:
        # fallback: prova prime 3 e 3 centrali
        family = sn[:3]
        model  = sn[5:8] if len(sn) >= 8 else ""
    return family, model


def _db_get_value_for(sn: str, key: str):
    """Legge ./database/<family>.xlsx e prende la colonna 'key' per il model code."""
    import pandas as pd, os, re
    fam, mod = _parse_sn_family_model(sn)
    if not fam or not mod:
        return None
    db_path = os.path.join("./database", f"{fam}.xlsx")
    if not os.path.isfile(db_path):
        return None
    try:
        df = pd.read_excel(db_path, dtype={"Unnamed: 0": str}, keep_default_na=False)
        # normalizza chiave modello
        df["Unnamed: 0"] = (df["Unnamed: 0"].astype(str).str.strip()
                            .str.replace(r"\.0+$", "", regex=True).str.zfill(3))
        key_mod = str(mod).zfill(3)
        row = df.loc[df["Unnamed: 0"] == key_mod]
        if isinstance(key, (list, tuple)):
            keys = list(key)
        elif isinstance(key, str):
            keys = [k.strip() for k in key.split("|") if k.strip()]
        else:
            keys = [str(key)]
        for k in keys:
            if (not row.empty) and (k in df.columns):
                val = row.iloc[0][k]
                try:
                    return float(str(val).replace(",", "."))
                except:
                    pass
        return None
    except Exception:
        return None


def _first_existing(*paths):
    for p in paths:
        if p and os.path.isfile(p):
            return p
    return None


def _default_logo():
    for ext in ("*.png","*.jpg","*.jpeg"):
        cand = glob.glob(os.path.join("misc","logo",ext))
        if cand: return cand[0]
    return None


def _default_header_footer():
    imgs = glob.glob(os.path.join("misc","carta intestata","*.png"))
    header = footer = None
    if imgs:
        for p in imgs:
            n = os.path.basename(p).lower()
            if ("head" in n or "header" in n) and not header: header = p
            if ("foot" in n or "footer" in n) and not footer: footer = p
        if not header: header = imgs[0]
        if not footer: footer = imgs[-1 if len(imgs)>1 else 0]
    return header, footer


def _load_template_html():
    p = os.path.join("templates","report_template.html")
    if os.path.isfile(p):
        return open(p,"r",encoding="utf-8").read()
    # Fallback HTML minimale (Bootstrapless, CSS inline)
    return """<!doctype html>
<html lang="it">
<head>
<meta charset="utf-8">
<title>{{ title }}</title>
<style>
@page { margin: 14mm 16mm; }                        /* margini pagina */
body { margin: 0; font-family: Arial, sans-serif; font-size: 12pt; color:#111; }
.header-hero { width:100%; max-height: 30mm; object-fit: contain; margin-bottom: 6mm; }
.center { text-align:center; }
.logo-hero   { height: 50mm; }               /* logo centrale PIÙ grande */
.meta        { color:#444; }

/* PRIMA PAGINA: tutto centrato verticalmente */
.first {
  position: relative;
  min-height: 260mm;                      /* area utile di una A4 con margini */
  display: flex; flex-direction: column;
  align-items: center; justify-content: center;
  text-align: center;
  padding-top: 22mm;                      /* riserva spazio per la barra in alto */
}
/* barra loghi solo nel frontespizio, attaccata in alto */
.front-top {
  position: absolute; top: 0; left: 0; right: 0;
  height: 20mm;
  display: flex; align-items: center; justify-content: center;
  padding: 0 10mm;
}
.front-top img {
  max-width: 100%;
  max-height: 100%;
  object-fit: contain;
}
.section { margin: 18px 0; }
table { width:100%; border-collapse: collapse; }
th, td { border:1px solid #ddd; padding:8px 10px; }
th { background: #f5f5f5; text-align:left; }
.img-full { width:100%; }


</style>
</head>
<body>
  <!-- PRIMA PAGINA: header grande + logo + titoli, tutto centrato -->
  <section class="first">
    {% if header_filename %}
    <div class="front-top">
      <img src="{{ header_filename }}" alt="Header">
    </div>
    {% endif %}
    
    {% if logo_filename %}<img class="logo-hero" src="{{ logo_filename }}">{% endif %}
    <div class="meta" style="margin-top:4mm; font-size:14pt;">
      <strong>{{ company or "GID Lab" }}</strong> – {{ date }}
    </div>
    <h1 style="margin-top:6mm;">{{ title }}</h1>

    <div class="meta" style="margin-top:6mm;">
      <div>Template: <span class="muted">{{ template_name }}</span></div>
      <div>Seriale UUT: <strong>{{ uut_serial }}</strong></div>
    </div>
  </section>

  <!-- PAGINA 2 -->
  <div class="page-break"></div>

  <main class="content">
    <!-- Risultato del test -->
    <section>
      <h2>Risultato del test</h2>
      <table>
        <thead>
          <tr>
            <th>Nome Test</th>
            <th>Risultato Test</th>
            <th>Valore di riferimento</th>
            <th>Tolleranza</th>
            <th>Valore misurato</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>{{ test_name }}</td>
            <td><strong>{{ result_text }}</strong></td>
            <td>{{ ref_value_fmt }}</td>
            <td>{{ tol_fmt }}</td>
            <td>{{ meas_value_fmt }}</td>
          </tr>
        </tbody>
      </table>
    </section>

    <section class="section">
  <h2>Dati allegati</h2>
  <div>File log: <code>{{ log_xlsx_basename }}</code></div>
  {% if graphs and graphs|length>0 %}
    <div style="margin-top:6px;">Grafici (Pdcx vs Vdcx):</div>
    <ul>
      {% for g in graphs %}
        <li>{{ g.ch }}{% if g.html %} — <a href="{{ g.html|basename }}">interattivo</a>{% endif %}</li>
      {% endfor %}
    </ul>
    {% for g in graphs %}
      {% if g.png %}<img src="{{ g.png|basename }}" class="img-full" style="margin-top:10px;">{% endif %}
    {% endfor %}
  {% endif %}
</section>



    <!-- Allarmi -->
    <section class="section">
      <h2>Allarmi</h2>
      {% if alarms|length>0 %}
      <table>
        <thead><tr><th>Timestamp</th><th>Codice</th><th>Fonte</th></tr></thead>
        <tbody>
          {% for a in alarms %}
            <tr><td>{{ a.ts }}</td><td>{{ a.code }}</td><td>{{ a.src }}</td></tr>
          {% endfor %}
        </tbody>
      </table>
      {% else %}
        <div class="muted">Nessun allarme disponibile.</div>
      {% endif %}
    </section>
  </main>
</body>
</html>
    """


def _find_vdc_columns(df: pd.DataFrame):
    return [c for c in df.columns if c.lower().startswith("voltage dc") and "[v" in c.lower()]


def _find_pout_column(df: pd.DataFrame):
    for c in df.columns:
        cl = c.lower()
        if "active output power" in cl and "[W" in cl:
            return c
    for c in df.columns:
        if c.lower().startswith("power dc") and "[W" in c.lower():
            return c
    return None


def _detect_dc_channels(df, allowed=None):
    """ Ritorna lista canali presenti ['DC1','DC2','DC3'] in base alle colonne. """
    chans = []
    for ch in ("DC1","DC2","DC3"):
        if allowed and ch not in allowed:
            continue
        v = f"Voltage {ch} [V]"
        p = f"Power {ch} [W]"
        if v in df.columns and p in df.columns:
            chans.append(f"{ch}")
    return chans


def _graphs_multi(df, out_base, allowed=None):
    """
    Crea (per ogni canale DCx presente) un PNG Matplotlib Pdcx vs Vdcx
    e un HTML Plotly interattivo (opzionale).
    Ritorna: list of {'ch':'DC1','png':'...', 'html':'...'}
    """
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import plotly.express as px

    items = []
    chans = _detect_dc_channels(df, allowed=allowed)
    for ch in chans:
        vcol = f"Voltage {ch} [V]"
        pcol = f"Power {ch} [W]"
        png = f"{out_base}_{ch}.png"
        html = f"{out_base}_{ch}.html"

        # PNG (Matplotlib)
        try:
            x = df[vcol].values
            y = df[pcol].values
            plt.figure()
            plt.scatter(x, y, s=12)
            plt.title(f"{pcol} vs {vcol}")
            plt.xlabel(vcol); plt.ylabel(pcol)
            plt.grid(True, alpha=0.3)
            plt.savefig(png, dpi=160, bbox_inches="tight")
            plt.close()
            if not os.path.isfile(png):
                png = ""
        except Exception as e:
            print(f"[WARN] Matplotlib PNG {ch} fallito: {e}")
            png = ""

        # HTML (Plotly) — non blocca se fallisce
        try:
            fig = px.scatter(df, x=vcol, y=pcol, title=f"{pcol} vs {vcol}")
            fig.write_html(html)
        except Exception as e:
            print(f"[WARN] Plotly HTML {ch} fallito: {e}")
            html = ""

        items.append({"ch": ch, "png": png, "html": html})
    return items


def _write_graphs(df, template_path, out_base_noext):
    base = os.path.basename(template_path).upper()
    vdc_cols = _find_vdc_columns(df)
    pout_col = _find_pout_column(df)

    preferred = None
    for tag in ("DC1","DC2","DC3"):
        if tag in base:
            for c in vdc_cols:
                if tag in c.upper():
                    preferred = c; break
    x_col = preferred or (vdc_cols[0] if vdc_cols else None)

    graph_html = out_base_noext + "_graph.html"
    graph_png = out_base_noext + "_graph.png"
    if x_col and pout_col:
        fig = px.scatter(df, x=x_col, y=pout_col, title=f"{pout_col} vs {x_col}")
        fig.write_html(graph_html)
        png_ok = False
        try:
            import matplotlib
            matplotlib.use("Agg")
            import matplotlib.pyplot as plt
            plt.figure()
            plt.scatter(df[x_col].values, df[pout_col].values)
            plt.title(f"{pout_col} vs {x_col}")
            plt.xlabel(x_col)
            plt.ylabel(pout_col)
            plt.grid(True, alpha=0.3)
            plt.savefig(graph_png, dpi=160, bbox_inches="tight")
            plt.close()
            png_ok = os.path.isfile(graph_png)
        except Exception as e2:
            print(f"[WARN] Fallback Matplotlib fallita: {e2}")
            graph_png = ""

        # normalizza: se nessun exporter ha prodotto un file valido, svuota il path
        if not png_ok or not os.path.isfile(graph_png):
            graph_png = ""
    else:
        graph_html = ""
        graph_png = ""
    return graph_html, graph_png


def _to_pdf_via_browser(html_path: str, out_pdf_path: str, timeout_s: int = 60) -> bool:
    """
    Converte HTML in PDF usando Edge/Chrome in headless.
    - usa lista di argomenti (niente shell=True)
    - usa Path.as_uri() per avere 'file:///C:/...'
    """
    from pathlib import Path
    candidates = [
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    exe = _first_existing(*candidates)
    if not exe:
        return False
    uri = Path(html_path).resolve().as_uri()  # -> file:///C:/path/file.html
    args = [exe, "--headless", f"--print-to-pdf={out_pdf_path}",
            "--print-to-pdf-no-header", "--no-pdf-header-footer",
            uri]
    if "chrome" in exe.lower():
        args[1] = "--headless=new"
    try:
        subprocess.run(args, shell=False, check=True, timeout=timeout_s)
        return os.path.isfile(out_pdf_path)
    except Exception:
        return False


def render_mppt_report_html(log_xlsx_path: str,
                            inverter_serials: list,
                            template_path: str,
                            out_html_path: str,
                            out_pdf_path: str or None = None,
                            logo_path: str or None = None,
                            header_path: str or None = None,
                            footer_path: str or None = None,
                            meta: dict or None = None):
    """Genera un report HTML (e opzionalmente PDF via Edge/Chrome) per 'curva MPPT'."""
    meta = meta or {}
    # Template Excel (per tabella di riepilogo)
    df_tpl = pd.read_excel(template_path)
    template_rows = []
    for col in df_tpl.columns[:12]:
        try:
            template_rows.append({"col": str(col), "val": str(df_tpl[col].iloc[0])})
        except Exception:
            pass

    # Log XLSX
    xls = pd.ExcelFile(log_xlsx_path, engine="openpyxl")
    # scegli foglio del primo seriale disponibile
    main_serial = inverter_serials[0] if inverter_serials else "INV1"
    sheet_name = None
    for sn in inverter_serials:
        if sn in xls.sheet_names:
            sheet_name = sn
            main_serial = sn
            break
    sheet_name = sheet_name or xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_name)

    # # ---- Calcolo PASS/FAIL in base al template ----

    # ---- Calcolo PASS/FAIL (PCH ± tol%) ----
    spec = _guess_spec_from_template(template_path)
    test_name = os.path.splitext(os.path.basename(template_path))[0]
    title = spec.get("title") or f"Report Test – {test_name}"

    if spec.get("mode") == "battery_cycle":
        tol_pct = float(spec.get("tol_percent", 5.0))
        pbat_col = spec.get("pbat_col") or _find_batt_power_col(df)
        ac_col = spec.get("ac_col", "Active Output Power [W]")
        settle_min_s = float(spec.get("settle_min_s", 5.0))
        settle_ratio = float(spec.get("settle_ratio", 0.2))

        # time axis
        t_s, _ = _build_time_axis(df)
        steps = _read_template_steps(template_path, main_serial)

        result_rows = []  # <-- nuova tabella multi-riga
        overall_ok = True

        if not pbat_col or pbat_col not in df.columns:
            result_rows = [{"name": "ciclo batteria", "result": "NON VALUTABILE",
                            "ref": "n/d", "tol": f"{tol_pct:.1f}%", "meas": "n/d"}]
            overall_ok = False
        else:
            import numpy as np
            s = pd.to_numeric(df[pbat_col], errors="coerce").to_numpy()
            # mediana di ciascuno step (positivi = carica, negativi = scarica)
            med_charge = [];
            med_discharge = []
            tcur = float(t_s[0] if len(t_s) else 0.0)
            # ---- Calcolo PASS/FAIL (PCH ± tol%) ----
            spec = _guess_spec_from_template(template_path)
            test_name = os.path.splitext(os.path.basename(template_path))[0]
            title = spec.get("title") or f"Report Test – {test_name}"

            if spec.get("mode") == "battery_cycle":
                tol_pct = float(spec.get("tol_percent", 5.0))
                pbat_col = spec.get("pbat_col") or _find_batt_power_col(df)
                ac_col = spec.get("ac_col", "Active Output Power [W]")
                settle_min_s = float(spec.get("settle_min_s", 5.0))
                settle_ratio = float(spec.get("settle_ratio", 0.2))

                # time axis
                t_s, _ = _build_time_axis(df)
                steps = _read_template_steps(template_path, main_serial)

                result_rows = []  # <-- nuova tabella multi-riga
                overall_ok = True

                if not pbat_col or pbat_col not in df.columns:
                    result_rows = [{"name": "ciclo batteria", "result": "NON VALUTABILE",
                                    "ref": "n/d", "tol": f"{tol_pct:.1f}%", "meas": "n/d"}]
                    overall_ok = False
                else:
                    import numpy as np
                    s = pd.to_numeric(df[pbat_col], errors="coerce").to_numpy()
                    # mediana di ciascuno step (positivi = carica, negativi = scarica)
                    med_charge = [];
                    med_discharge = []
                    tcur = float(t_s[0] if len(t_s) else 0.0)
                    for st in steps:
                        dur = float(st["dur"]);
                        psp = float(st["psp"])
                        if dur <= 0:
                            tcur += dur;
                            continue
                        t1 = tcur + dur
                        mask = (t_s >= tcur) & (t_s <= t1)
                        if not np.any(mask):
                            tcur = t1;
                            continue
                        settle = max(settle_min_s, settle_ratio * dur)
                        mask2 = (t_s >= (tcur + settle)) & (t_s <= t1)
                        if not np.any(mask2): mask2 = mask
                        ms = s[mask2];
                        ms = ms[np.isfinite(ms)]
                        if ms.size:
                            med = float(np.median(ms))
                            if psp >= 0:
                                med_charge.append(med)
                            else:
                                med_discharge.append(med)
                        tcur = t1

                    # riferimenti dai DB
                    pbat_db = _db_get_value_for(main_serial, "P BAT") or 0.0
                    pbat_unit = spec.get("pbat_unit") or "W"
                    ref_base = float(pbat_db) / _unit_scale(pbat_unit, "W")
                    ref_charge = ref_base * 1.05
                    ref_discharge = - ref_base * 1.05

                    # mediana aggregata per modalità
                    m_charge = (float(np.median(med_charge)) if med_charge else None)
                    m_discharge = (float(np.median(med_discharge)) if med_discharge else None)

                    # check: carica <= ref_charge ; scarica >= ref_discharge
                    def fmt(v):
                        return _fmt_value(v, "W")

                    if m_charge is not None:
                        ok = (m_charge <= ref_charge + 1e-6)
                        overall_ok &= ok
                        result_rows.append({
                            "name": "ciclo batteria (carica)",
                            "result": ("PASS" if ok else "FAIL"),
                            "ref": _fmt_value(ref_charge, "W"),
                            "tol": "+5.0%",  # incorporata nel ref
                            "meas": fmt(m_charge),
                        })
                    if m_discharge is not None:
                        ok = (m_discharge >= ref_discharge - 1e-6)
                        overall_ok &= ok
                        result_rows.append({
                            "name": "ciclo batteria (scarica)",
                            "result": ("PASS" if ok else "FAIL"),
                            "ref": _fmt_value(ref_discharge, "W"),
                            "tol": "+5.0%",  # incorporata nel ref
                            "meas": fmt(m_discharge),
                        })
                    if not result_rows:
                        result_rows = [{"name": "ciclo batteria", "result": "NON VALUTABILE",
                                        "ref": "n/d", "tol": "n/d", "meas": "n/d"}]
                        overall_ok = False

            # riferimenti dai DB
            # pbat_db = _db_get_value_for(main_serial, "P BAT") or 0.0
            # pbat_unit = spec.get("pbat_unit") or "W"
            # ref_base = float(pbat_db) #* _unit_scale(pbat_unit, "W")
            # ref_charge = ref_base * 1.05
            # ref_discharge = - ref_base * 1.05

            # mediana aggregata per modalità
            # m_charge = (float(np.median(med_charge)) if med_charge else None)
            # m_discharge = (float(np.median(med_discharge)) if med_discharge else None)
            #
            # # check: carica <= ref_charge ; scarica >= ref_discharge
            # def fmt(v): return _fmt_value(v, "W")
            #
            # if m_charge is not None:
            #     ok = (m_charge <= ref_charge + 1e-6)
            #     overall_ok &= ok
            #     result_rows.append({
            #         "name": "ciclo batteria (carica)",
            #         "result": ("PASS" if ok else "FAIL"),
            #         "ref": _fmt_value(ref_charge, "W"),
            #         "tol": "+5.0%",  # incorporata nel ref
            #         "meas": fmt(m_charge),
            #     })
            # if m_discharge is not None:
            #     ok = (m_discharge >= ref_discharge - 1e-6)
            #     overall_ok &= ok
            #     result_rows.append({
            #         "name": "ciclo batteria (scarica)",
            #         "result": ("PASS" if ok else "FAIL"),
            #         "ref": _fmt_value(ref_discharge, "W"),
            #         "tol": "+5.0%",  # incorporata nel ref
            #         "meas": fmt(m_discharge),
            #     })
            # if not result_rows:
            #     result_rows = [{"name": "ciclo batteria", "result": "NON VALUTABILE",
            #                     "ref": "n/d", "tol": "n/d", "meas": "n/d"}]
            #     overall_ok = False

        # testo intestazione (compatibilità col template)
        result_text = "PASS" if overall_ok else "FAIL"
        ref_value_fmt = result_rows[0]['ref']
        tol_fmt = result_rows[0]['tol']
        meas_value_fmt = result_rows[0]['meas']

    elif spec.get("mode") == "curva_mppt":
        # ----- MPPT: check per singolo canale DCx -----
        ref_value = None
        tol_pct = spec.get("tol_percent", 5.0)
        ref_unit = spec.get("ref_unit", 'kW')
        meas_unit = spec.get("meas_unit", 'W')
        scale_rm = _unit_scale(ref_unit, meas_unit)
        if spec.get("ref_db_key"):
            ref_value = _db_get_value_for(main_serial, spec["ref_db_key"])
        allowed = _allowed_chans_from_template(template_path)
        dc_chans = _detect_dc_channels(df, allowed=allowed)
        result_rows = []
        overall_ok = True
        for ch in dc_chans:
            col = f"Power {ch} [W]"
            if (col in df.columns) and (ref_value is not None):
                ser = pd.to_numeric(df[col], errors="coerce")
                meas = float(ser.max(skipna=True)) if ser.notna().any() else None
                meas = meas * scale_rm
                if meas is None:
                    result_rows.append({"name": f"curva MPPT {ch}",
                                        "result": "NON VALUTABILE",
                                        "ref": "n/d", "tol": f"{tol_pct:.1f}%", "meas": "n/d"})
                    overall_ok = False
                    continue
                ref_m = ref_value * scale_rm
                lo = ref_m * (1 - tol_pct / 100.0)
                hi = ref_m * (1 + tol_pct / 100.0)
                ok = (lo <= meas <= hi)
                overall_ok &= ok
                result_rows.append({"name": f"curva MPPT {ch}",
                                    "result": ("PASS" if ok else "FAIL"),
                                    "ref": _fmt_value(ref_m, meas_unit or "W"),
                                    "tol": f"{tol_pct:.1f}%",
                                    "meas": _fmt_value(meas, meas_unit or "W")
                                    })
        if not result_rows:
            result_rows = [
                {"name": "curva MPPT", "result": "NON VALUTABILE", "ref": "n/d", "tol": "n/d", "meas": "n/d"}]
            overall_ok = False

        result_text = "PASS" if overall_ok else "FAIL"
        ref_value_fmt = result_rows[0]['ref']
        tol_fmt = result_rows[0]['tol']
        meas_value_fmt = result_rows[0]['meas']
    # grafici
    base_out = os.path.splitext(out_html_path)[0]
    graphs = []
    allowed = _allowed_chans_from_template(template_path)
    # Se è ciclo batteria -> SOLO PBAT vs tempo; altrimenti i classici Pdcx vs Vdcx
    if spec.get("mode") == "battery_cycle":
        try:
            import matplotlib
            matplotlib.use("Agg")
            import matplotlib.pyplot as plt
            import plotly.express as px
            pcol = _find_batt_power_col(df)
            if pcol:
                png_p = f"{base_out}_PBAT_ts.png"
                html_p = f"{base_out}_PBAT_ts.html"
                plt.figure()
                y = pd.to_numeric(df[pcol], errors="coerce")
                plt.plot(y.index, y.values)
                plt.title(f"{pcol} vs samples");
                plt.xlabel("samples");
                plt.ylabel(pcol)
                plt.grid(True, alpha=0.3)
                plt.savefig(png_p, dpi=160, bbox_inches="tight");
                plt.close()
                try:
                    figp = px.line(df, y=pcol, title=f"{pcol} vs time")
                    figp.write_html(html_p)
                except Exception:
                    html_p = ""
                graphs.append({"ch": "PBAT", "png": png_p if os.path.isfile(png_p) else "", "html": html_p})
        except Exception as e:
            print(f"[WARN] grafico PBAT fallito: {e}")
    elif spec.get("mode") == "curva_mppt":
        graphs = _graphs_multi(df, base_out, allowed=allowed)

    # Allarmi allegati (merge *_LogErrori)
    alarms = []
    for sn in inverter_serials:
        sname = f"{sn}_LogErrori"
        if sname in xls.sheet_names:
            dfa = pd.read_excel(xls, sheet_name=sname)
            for _, r in dfa.iterrows():
                alarms.append({
                    "ts": str(r.get("timestamp","")),
                    "code": str(r.get("code_hex", r.get("code_dec",""))),
                    "src": str(r.get("source","HIST"))
                })

    # Default immagini se mancanti
    if not logo_path: logo_path = _default_logo()
    if not header_path or not footer_path:
        hdef, fdef = _default_header_footer()
        header_path = header_path or hdef
        footer_path = footer_path or fdef

    # Prepara cartella di lavoro per riferimenti immagini relativi
    workdir = os.path.dirname(os.path.abspath(out_html_path))
    assets = []
    decor_assets = []

    def _stage_asset(pathlike: str or None):
        if not pathlike:
            return
        try:
            src = os.path.abspath(pathlike)
            if os.path.isfile(src):
                dst = os.path.join(workdir, os.path.basename(src))
                if os.path.abspath(src) != os.path.abspath(dst):
                    try:
                        shutil.copy(src, dst)
                    except Exception:
                        pass
                assets.append(dst)
        except Exception:
            pass

    # logo/header (footer ormai non lo usiamo, ma tenerlo non fa danni)
    for p in (logo_path, header_path, footer_path):
        before = len(assets)
        _stage_asset(p)
        if len(assets) > before:
            decor_assets.append(assets[-1])
    # grafici Pdcx vs Vdcx (png + html interattivo)
    for it in (graphs or []):
        _stage_asset(it.get("png"))
        _stage_asset(it.get("html"))

    # Render HTML
    tpl_text = _load_template_html()

    env = Environment()
    env.filters['basename'] = lambda p: (os.path.basename(p) if p else "")
    html = env.from_string(tpl_text).render(
        title=title,  # nuovo titolo
        subtitle="",  # non usiamo più
        author=meta.get("author", ""),
        company=meta.get("company", "GID Lab"),
        template_name=os.path.basename(template_path),
        uut_serial=main_serial,  # "Seriale UUT"
        date=pd.Timestamp.now().strftime("%Y-%m-%d"),
        # rimosso: template_rows
        result_text=result_text,
        test_name=test_name,
        ref_value_fmt=ref_value_fmt,
        tol_fmt=tol_fmt,
        meas_value_fmt=meas_value_fmt,
        log_xlsx_basename=os.path.basename(log_xlsx_path),
        graphs=[{"ch": it["ch"],
                 "png": os.path.basename(it["png"]) if it.get("png") else "",
                 "html": os.path.basename(it["html"]) if it.get("html") else ""} for it in graphs],
        logo_filename=os.path.basename(logo_path) if logo_path else "",
        header_filename=os.path.basename(header_path) if header_path else "",
        footer_filename="",
        alarms=alarms
    )
    with open(out_html_path,"w",encoding="utf-8") as f:
        f.write(html)

    # PDF: proviamo, ma restituiamo sempre l'HTML
    pdf_ok = False
    if out_pdf_path:
        # # 1) WeasyPrint
        # try:
        #     from weasyprint import HTML
        #     HTML(filename=out_html_path).write_pdf(out_pdf_path)
        #     pdf_ok = True
        # except Exception as e:
        #     print(f"[WARN] WeasyPrint PDF fallito: {e}")
        #
        # # 2) pdfkit / wkhtmltopdf
        # if not pdf_ok:
        #     try:
        #         import pdfkit
        #         pdfkit.from_file(out_html_path, out_pdf_path)
        #         pdf_ok = True
        #     except Exception as e2:
        #         print(f"[WARN] pdfkit PDF fallito: {e2}")

        # 3) Edge/Chrome headless
        if not pdf_ok:
            ok, err = _headless_pdf(out_html_path, out_pdf_path)
            pdf_ok = ok
            if not ok:
                print(f"[WARN] Headless PDF fallito: {err}")
    return out_html_path, (out_pdf_path if pdf_ok else None), graphs


def _compute_result_for(log_xlsx_path, template_path, serial):
    """Riuso della stessa logica di valutazione (AC / ΣDC) per l'index."""
    import pandas as pd
    try:
        xls = pd.ExcelFile(log_xlsx_path, engine="openpyxl")
        # scegli foglio del seriale, altrimenti il primo
        sheet = serial if serial in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet)
    except Exception:
        return "n/d"

    spec = _guess_spec_from_template(template_path)
    ref = _db_get_value_for(serial, spec.get("ref_db_key") or "PCH")
    tol = spec.get("tol_percent") or 5.0

    def _in(val, ref, tol):
        lo = ref*(1-tol/100); hi = ref*(1+tol/100); return lo <= val <= hi

    # AC
    ac = None
    if "Active Output Power [W]" in df.columns:
        s = pd.to_numeric(df["Active Output Power [W]"], errors="coerce")
        if s.notna().any(): ac = float(s.max(skipna=True))

    # ΣDC
    dc_chans = _detect_dc_channels(df)
    dcs = None
    if dc_chans:
        ss = 0.0; ok = False
        for ch in dc_chans:
            c = f"Power {ch} [W]"
            if c in df.columns:
                ser = pd.to_numeric(df[c], errors="coerce")
                if ser.notna().any():
                    ss += float(ser.max(skipna=True)); ok = True
        if ok: dcs = ss

    checks = []
    if ref is not None:
        if ac is not None:  checks.append(_in(ac, ref, tol))
        if dcs is not None: checks.append(_in(dcs, ref, tol))
    if checks:
        return "PASS" if all(checks) else "FAIL"
    return "NON VALUTABILE"

def render_session_index(session_dir: str, out_html_path: str or None = None, out_pdf_path: str or None = None,
                         company: str = "Lab", logo_path: str or None = None,
                         header_path: str or None = None, footer_path: str or None = None):
    from jinja2 import Template
    session_dir = os.path.abspath(session_dir)
    if out_html_path is None: out_html_path = os.path.join(session_dir, "index.html")
    if out_pdf_path  is None: out_pdf_path  = os.path.join(session_dir, "index.pdf")

    # raccogli test: coppie (xlsx, pdf/html)
    files = os.listdir(session_dir)
    xlsx = [f for f in files if f.lower().endswith(".xlsx")]
    pdfs = {os.path.splitext(f)[0]: f for f in files
            if f.lower().endswith(".pdf") and f.lower().startswith("report_")}
    htmls = {os.path.splitext(f)[0]: f for f in files
             if f.lower().endswith(".html") and f.lower().startswith("report_")}

    # anche i file nella gemella ./Reports/<sessione>, con link RELATIVI all’index
    data_dir = os.path.abspath(os.path.dirname(session_dir))  # .../<root>/Data
    root_dir = os.path.abspath(os.path.dirname(data_dir))  # .../<root>
    reports_dir = os.path.join(root_dir, "Reports", os.path.basename(session_dir))
    if os.path.isdir(reports_dir):
        for f in os.listdir(reports_dir):
            stem, ext = os.path.splitext(f)
            if not f.lower().startswith("report_"): continue
            if ext.lower() not in (".pdf", ".html"): continue
            rel = os.path.relpath(os.path.join(reports_dir, f), start=session_dir)
            if ext.lower()==".pdf":  pdfs[stem]  = rel
            if ext.lower()==".html": htmls[stem] = rel

    # SN di default dal nome directory
    sess_base = os.path.basename(session_dir)
    default_sn = sess_base.split("_")[0] if "_" in sess_base else "INV"

    rows = []
    for fx in sorted(xlsx):
        test_name = os.path.splitext(fx)[0]
        # skip "custom"
        if test_name.strip().lower() == "custom":
            continue
        # prova a trovare template omonimo
        tpl = os.path.join(".", "template", f"{test_name}.xlsx")
        if not os.path.isfile(tpl):
            continue
        # deduci seriale dal workbook o usa default
        try:
            import pandas as pd
            xl = pd.ExcelFile(os.path.join(session_dir, fx), engine="openpyxl")
            serials = [s for s in xl.sheet_names if s and "_LogErrori" not in s]
            sn = serials[0] if serials else default_sn
        except Exception:
            sn = default_sn

        result = _compute_result_for(os.path.join(session_dir, fx), tpl, sn)
        base = f"Report_{test_name}_{sn}"
        link = pdfs.get(base) or htmls.get(base) or ""
        rows.append({"test": test_name, "result": result, "link": link})

    # HTML
    tpl_html = """<!doctype html><html lang="it"><head><meta charset="utf-8">
<title>Indice Sessione</title>
<style>
@page { margin: 14mm 16mm; }
body { font-family: Arial, sans-serif; font-size: 12pt; color:#111; margin:0; }
.brand { display:flex; align-items:center; gap:12px; margin:10mm 0 4mm 0; }
.brand img.logo { height: 28px; }
.muted { color:#666; }
table { width:100%; border-collapse: collapse; }
th, td { border:1px solid #ddd; padding:8px 10px; }
th { background:#f5f5f5; text-align:left; }
.pass { color:#0a8a0a; font-weight:bold; }
.fail { color:#c21717; font-weight:bold; }
</style></head><body>
<div class="brand">
  {% if logo_filename %}<img class="logo" src="{{ logo_filename }}">{% endif %}
  <div><strong>{{ company }}</strong><br><span class="muted">{{ session_name }}</span></div>
</div>
<h1>Indice sessione</h1>
<table>
  <thead><tr><th>Nome test</th><th>Risultato</th><th>Report</th></tr></thead>
  <tbody>
  {% for r in rows %}
    <tr>
      <td>{{ r.test }}</td>
      <td class="{{ 'pass' if r.result=='PASS' else ('fail' if r.result=='FAIL' else '') }}">{{ r.result }}</td>
      <td>{% if r.link %}<a href="{{ r.link }}">{{ r.link }}</a>{% else %}<span class="muted">n/d</span>{% endif %}</td>
    </tr>
  {% endfor %}
  </tbody>
</table>
</body></html>"""
    html = Template(tpl_html).render(
        company=company,
        session_name=os.path.basename(session_dir),
        rows=rows,
        logo_filename=os.path.basename(logo_path) if logo_path else ""
    )
    with open(out_html_path, "w", encoding="utf-8") as f:
        f.write(html)

    # prova PDF dell'index (non obbligatorio)
    pdf_ok = _to_pdf_via_browser(out_html_path, out_pdf_path)
    return out_html_path, (out_pdf_path if pdf_ok else None)

