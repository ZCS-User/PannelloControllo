# drivers/report.py
import os, datetime, tempfile, shutil, subprocess
import pandas as pd
from jinja2 import Template
import plotly.express as px
import glob

def _load_template_text():
    tpl_path = os.path.join("templates", "report_template.tex")
    if os.path.isfile(tpl_path):
        return open(tpl_path, "r", encoding="utf-8").read()
    # fallback minimo
    return r"""\documentclass[11pt,a4paper]{report}
\usepackage[margin=2.5cm]{geometry}
\usepackage{graphicx}
\usepackage{fancyhdr}
\usepackage{booktabs}
\usepackage{longtable}
\usepackage{hyperref}
\pagestyle{fancy}\fancyhf{}
\lhead{\includegraphics[height=1.2cm]{ {{ logo_filename }} }}\rhead{ {{ company|default("ZCS") }} }
\cfoot{\thepage}
\begin{document}
\begin{titlepage}\centering\vspace*{2cm}
{\Huge \textbf{ {{ title }} }\par}\vspace{0.8cm}
{\Large {{ subtitle|default("") }}\par}\vspace{1.2cm}
{\large \textbf{Seriale principale:} {{ main_serial }}\par}\vspace{0.6cm}
{\large \textbf{Data:} {{ date }}\par}\vfill
{\large \textbf{Autore:} {{ author|default("") }}\par}
\end{titlepage}
\tableofcontents\newpage
\chapter*{Introduzione}
Template: \texttt{ {{ template_name }} }\\
Seriale: \texttt{ {{ main_serial }} }
\section*{Righe template}
\begin{longtable}{@{}{{ '{' }}}p{0.4\textwidth}p{0.55\textwidth}{{ '}' }}@{}}
\toprule \textbf{Colonna} & \textbf{Valore} \\ \midrule
{% for row in template_rows %} {{row.col}} & {{row.val}}\\ {% endfor %}
\bottomrule
\end{longtable}
\chapter*{Risultato del test}
{{ result_text }}
\chapter*{Dati allegati}
File log: \texttt{ {{ log_xlsx_basename }} } \\
Grafico: \href{ {{ graph_html_link }} }{versione interattiva}
\begin{center}\includegraphics[width=\textwidth]{ {{ graph_png_filename }} }\end{center}
\chapter*{Allarmi}
{% if alarms|length>0 %}
\begin{longtable}{@{}lll@{}}
\toprule Timestamp & Codice & Fonte \\ \midrule
{% for a in alarms %} {{a.ts}} & {{a.code}} & {{a.src}} \\
{% endfor %}\bottomrule
\end{longtable}
{% else %} Nessun allarme disponibile (placeholder). {% endif %}
\end{document}"""

def _first_existing(*paths):
    for p in paths:
        if p and os.path.isfile(p):
            return p
    return None

def _default_logo():
    # .\misc\logo\  → prendi il primo .png/.jpg
    for ext in ("*.png","*.jpg","*.jpeg"):
        cand = glob.glob(os.path.join("misc","logo",ext))
        if cand:
            return cand[0]
    return None

def _default_header_footer():
    # .\misc\carta intestata\ → cerca png plausibili (nomi con 'head'/'foot', altrimenti primi due)
    imgs = glob.glob(os.path.join("misc","carta intestata","*.png"))
    header = footer = None
    if imgs:
        for p in imgs:
            name = os.path.basename(p).lower()
            if ("head" in name or "header" in name) and not header:
                header = p
            if ("foot" in name or "footer" in name) and not footer:
                footer = p
        if not header:
            header = imgs[0]
        if not footer:
            footer = imgs[-1 if len(imgs)>1 else 0]
    return header, footer

def _find_vdc_columns(df: pd.DataFrame):
    # colonne tipiche: Voltage DC1 [V], Voltage DC2 [V], Voltage DC3 [V]
    cols = [c for c in df.columns if c.lower().startswith("voltage dc") and "[v" in c.lower()]
    return cols

def _find_pout_column(df: pd.DataFrame):
    # tipiche: Active Output Power [kW] oppure Power DCx [kW]
    for c in df.columns:
        cl = c.lower()
        if "active output power" in cl and "[kw" in cl:
            return c
    # fallback: Power DC1/2/3
    for c in df.columns:
        if c.lower().startswith("power dc") and "[kw" in c.lower():
            return c
    return None

def render_mppt_report(log_xlsx_path: str,
                       inverter_serials: list,
                       template_path: str,
                       out_pdf_path: str,
                       logo_path: str = None,
                       header_path: str = None,
                       footer_path: str = None,
                       meta: dict = None):
    """Genera report 'curva MPPT' singolo test."""
    meta = meta or {}
    # carica template Excel (righe da mostrare)
    df_tpl = pd.read_excel(template_path)
    template_rows = []
    # prendi tutte le colonne e prima riga per 'panoramica'
    for col in df_tpl.columns[:10]:  # limita
        try:
            v = df_tpl[col].iloc[0]
            template_rows.append({"col": str(col), "val": str(v)})
        except Exception:
            pass

    # carica log (foglio primo inverter)
    xls = pd.ExcelFile(log_xlsx_path, engine="openpyxl")
    # sceglie il foglio del *primo* seriale
    main_serial = inverter_serials[0] if inverter_serials else "INV1"
    sheet_name = None
    for sn in inverter_serials:
        if sn in xls.sheet_names:
            sheet_name = sn
            main_serial = sn
            break
    if not sheet_name:
        # fallback: primo sheet
        sheet_name = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_name)

    # trova Vdc_x presenti e colonna Pout
    vdc_cols = _find_vdc_columns(df)
    pout_col = _find_pout_column(df)
    # se ci sono più Vdc, scegli quella coerente con il nome template (DC1/DC2/DC3)
    preferred = None
    base = os.path.basename(template_path).upper()
    for tag in ("DC1","DC2","DC3"):
        if tag in base:
            for c in vdc_cols:
                if tag in c.upper():
                    preferred = c; break
    x_col = preferred or (vdc_cols[0] if vdc_cols else None)

    # genera grafico (PNG + HTML interattivo)
    workdir = os.path.dirname(out_pdf_path) or "."
    graph_png = os.path.join(workdir, f"{os.path.splitext(os.path.basename(out_pdf_path))[0]}_graph.png")
    graph_html = os.path.join(workdir, f"{os.path.splitext(os.path.basename(out_pdf_path))[0]}_graph.html")

    if x_col and pout_col:
        fig = px.scatter(df, x=x_col, y=pout_col, title=f"{pout_col} vs {x_col}")
        fig.write_html(graph_html)
        # PNG: usa static export plotly se disponibile, altrimenti un fallback rapido
        try:
            # richiede kaleido (pip install -U kaleido)
            fig.write_image(graph_png, scale=2)
        except Exception:
            # fallback rudimentale: salva CSV filtrato e l'utente può rigenerare
            df[[x_col, pout_col]].to_csv(graph_png.replace(".png",".csv"), index=False)
    else:
        graph_html = ""
        graph_png = ""

    # allarmi: unisci sheet *_LogErrori se esistono
    alarms = []
    for sn in inverter_serials:
        name = f"{sn}_LogErrori"
        if name in xls.sheet_names:
            dfa = pd.read_excel(xls, sheet_name=name)
            for _, r in dfa.iterrows():
                alarms.append({
                    "ts": str(r.get("timestamp","")),
                    "code": str(r.get("code_hex", r.get("code_dec",""))),
                    "src": str(r.get("source","HIST"))
                })
    # placeholder se vuoto
    if not alarms:
        alarms = []

    # context LaTeX
    ctx = {
        "title": meta.get("title","Report Test – Curva MPPT"),
        "subtitle": meta.get("subtitle",""),
        "author": meta.get("author",""),
        "company": meta.get("company",""),
        "template_name": os.path.basename(template_path),
        "main_serial": main_serial,
        "date": datetime.datetime.now().strftime("%Y-%m-%d"),
        "template_rows": template_rows,
        "result_text": "NON APPLICABILE",
        "log_xlsx_basename": os.path.basename(log_xlsx_path),
        "graph_png_filename": os.path.basename(graph_png) if graph_png else "",
        "graph_html_link": graph_html.replace("\\","/") if graph_html else "",
        "alarms": alarms
    }

    # prepara cartella temp e compila LaTeX
    tmp = tempfile.mkdtemp(prefix="rep_")
    try:
        tpl_text = _load_template_text()
        # tex = Template(tpl_text).render(logo_filename=os.path.basename(logo_path) if logo_path else "",
        #                                 **ctx)

        # default files if not passed
        if not logo_path:
            logo_path = _default_logo()
        if not header_path or not footer_path:
            hdef, fdef = _default_header_footer()
            header_path = header_path or hdef
            footer_path = footer_path or fdef

        tex = Template(tpl_text).render(
            logo_filename=os.path.basename(logo_path) if logo_path else "",
            header_filename=os.path.basename(header_path) if header_path else "",
            footer_filename=os.path.basename(footer_path) if footer_path else "",
            **ctx
        )
        # copia logo/grafico se presenti
        if logo_path and os.path.isfile(logo_path):
            shutil.copy(logo_path, os.path.join(tmp, os.path.basename(logo_path)))
        # copia header/footer se presenti
        if header_path and os.path.isfile(header_path):
            shutil.copy(header_path, os.path.join(tmp, os.path.basename(header_path)))
        if footer_path and os.path.isfile(footer_path):
            shutil.copy(footer_path, os.path.join(tmp, os.path.basename(footer_path)))
        if graph_png and os.path.isfile(graph_png):
            shutil.copy(graph_png, os.path.join(tmp, os.path.basename(graph_png)))

        tex_path = os.path.join(tmp, "report.tex")
        with open(tex_path, "w", encoding="utf-8") as f:
            f.write(tex)

        # compila
        for _ in range(2):
            subprocess.run(["pdflatex","-interaction=nonstopmode","-halt-on-error","report.tex"],
                           cwd=tmp, check=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        shutil.copy(os.path.join(tmp,"report.pdf"), out_pdf_path)
        return out_pdf_path, graph_html
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
