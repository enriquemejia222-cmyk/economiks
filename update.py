#!/usr/bin/env python3
"""Economiks: descarga todas las fuentes y genera data.json.

Uso:  python3 -I update.py --out DIR [--hist historial_riesgo.json]
Requiere: pip install --break-system-packages pymupdf openpyxl numpy
Fuentes: FMI WEO (DataMapper API), Banco Mundial API, BCRP API (EMBIG mensual),
         BCRD (EMBI.xlsx serie histórica diaria + EMBI.pdf con el dato del día).
El historial guarda el dato diario de cada país leído del PDF del BCRD, para
cubrir los meses que no traen las otras fuentes. Se lee y se reescribe en --out.
"""
import argparse, collections, datetime as dt, json, os, re, statistics, subprocess, sys, tempfile, urllib.request

C = {"ARG":"Argentina","BOL":"Bolivia","BRA":"Brasil","CHL":"Chile","COL":"Colombia","CRI":"Costa Rica",
     "DOM":"República Dominicana","ECU":"Ecuador","SLV":"El Salvador","GTM":"Guatemala","HND":"Honduras",
     "HTI":"Haití","MEX":"México","NIC":"Nicaragua","PAN":"Panamá","PRY":"Paraguay","PER":"Perú","URY":"Uruguay","VEN":"Venezuela"}
REF = {"USA":"Estados Unidos"}   # referencia: datos anuales sí, riesgo país no (es la base del índice)
ALL = {**C, **REF}
IMF_IND = ["NGDP_RPCH","PCPIPCH","LUR","GGXWDG_NGDP","GGXCNL_NGDP","BCA_NGDPD","NGDPDPC"]
UA = {"User-Agent":"Mozilla/5.0 (Economiks data updater)"}
XLS_NAMES = {"Global":"GLB","LATINO":"LAT","REP DOM":"DOM","Argentina":"ARG","Bolivia":"BOL","Brasil":"BRA","Chile":"CHL",
     "Colombia":"COL","Costa Rica":"CRI","Ecuador":"ECU","El Salvador":"SLV","Guatemala":"GTM","Honduras":"HND",
     "México":"MEX","Paraguay":"PRY","Perú":"PER","Panamá":"PAN","Uruguay":"URY","Venezuela":"VEN"}
PDF_NAMES = {"Argentina":"ARG","Ecuador":"ECU","Bolivia":"BOL","El Salvador":"SLV","Colombia":"COL","México":"MEX",
     "GLOBAL":"GLB","LATINO":"LAT","Honduras":"HND","REP DOM":"DOM","Brasil":"BRA","Costa Rica":"CRI","Guatemala":"GTM",
     "Panamá":"PAN","Paraguay":"PRY","Perú":"PER","Chile":"CHL","Uruguay":"URY","Venezuela":"VEN"}
BCRP_OLD = {"LAT":"PN01137XM","ARG":"PN01130XM","BRA":"PN01131XM","CHL":"PN01132XM","COL":"PN01133XM",
            "ECU":"PN01134XM","MEX":"PN01135XM","PER":"PN01129XM","VEN":"PN01136XM"}
MES_ES = {"enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,"julio":7,"agosto":8,"septiembre":9,"setiembre":9,"octubre":10,"noviembre":11,"diciembre":12}
BCRP_M = {"Ene":1,"Feb":2,"Mar":3,"Abr":4,"May":5,"Jun":6,"Jul":7,"Ago":8,"Sep":9,"Set":9,"Oct":10,"Nov":11,"Dic":12}

def get(url, timeout=90):
    last = None
    for headers in (UA, {}, {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}):
        for _ in range(2):
            try:
                with urllib.request.urlopen(urllib.request.Request(url, headers=headers), timeout=timeout) as r:
                    return r.read()
            except Exception as e:
                last = e
    raise RuntimeError(f"no se pudo descargar {url}: {last}")

def log(*a): print(*a, file=sys.stderr)

# ---------- FMI y Banco Mundial ----------
def fetch_imf():
    out, vintage = {}, None
    meta = json.loads(get("https://www.imf.org/external/datamapper/api/v1/indicators"))["indicators"]
    src = meta.get("NGDP_RPCH", {}).get("source", "")
    m = re.search(r"\(([^)]+)\)", src); vintage = m.group(1) if m else src
    for ind in IMF_IND:
        d = json.loads(get(f"https://www.imf.org/external/datamapper/api/v1/{ind}"))["values"][ind]
        out[ind] = {c: {str(int(y)): round(v, 2) for y, v in d.get(c, {}).items() if v is not None and 1990 <= int(y) <= 2031} for c in ALL}
    return out, vintage

def fetch_wb():
    url = ("https://api.worldbank.org/v2/country/" + ";".join(ALL) +
           "/indicator/BX.TRF.PWKR.DT.GD.ZS?format=json&per_page=5000&date=1990:2031")
    rows = json.loads(get(url))[1]
    rem = {c: {} for c in ALL}
    for x in rows:
        if x["countryiso3code"] in ALL and x["value"] is not None:
            rem[x["countryiso3code"]][x["date"]] = round(x["value"], 2)
    return rem

# ---------- Riesgo país ----------
def fetch_bcrp():
    res = {}
    end = dt.date.today().strftime("%Y-%m").replace("-0", "-")
    for k, code in BCRP_OLD.items():
        try:
            d = json.loads(get(f"https://estadisticas.bcrp.gob.pe/estadisticas/series/api/{code}/json/2006-1/{end}"))
            s = {}
            for p in d["periods"]:
                mo, yr = p["name"].split(".")
                try: s[f"{yr}-{BCRP_M[mo]:02d}"] = float(p["values"][0])
                except (ValueError, KeyError): pass
            res[k] = s
        except Exception as e:
            log("BCRP falló", k, e)
    return res

def parse_xlsx(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    hi = next(i for i, r in enumerate(rows) if r and r[0] == "Fecha")
    hdr = rows[hi]
    cols = {j: XLS_NAMES[h] for j, h in enumerate(hdr) if h in XLS_NAMES}
    daily, prev = [], None
    for r in rows[hi + 1:]:
        d = r[0]
        if isinstance(d, str):
            try: d = dt.datetime.strptime(d.strip(), "%d-%b-%y")
            except ValueError: d = None
        if not isinstance(d, dt.datetime): continue
        d = d.date()
        if prev and abs((d - prev).days) > 40: continue  # fecha mal tecleada en la fuente
        prev = d
        vals = {k: r[j] * 100 for j, k in cols.items() if isinstance(r[j], (int, float))}
        daily.append((d, vals))
    return daily

def parse_pdf(path):
    txt = subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True, text=True).stdout
    blk = txt.split("RIESGO-PAÍS DE ECONOMÍAS LATINOAMERICANAS")[1]
    m = re.search(r"(\d{1,2}) de ([a-záéíóú]+),? (\d{4})", blk)
    date = dt.date(int(m.group(3)), MES_ES[m.group(2).lower()], int(m.group(1)))
    blk2 = blk.split("Fuente")[0]
    snap = {}
    for line in blk2.splitlines():
        mm = re.match(r"\s*([A-Za-zÁÉÍÓÚáéíóúñ ]+?)\s{2,}(-?\d[\d,\.]*)\s*$", line)
        if mm and mm.group(1).strip() in PDF_NAMES:
            snap[PDF_NAMES[mm.group(1).strip()]] = float(mm.group(2).replace(",", ""))
    v = re.search(r"Venezuela se sitúa en ([\d,\.]+)", blk)
    if v: snap["VEN"] = float(v.group(1).replace(",", ""))
    return date, snap

def pdf_chart(path, daily, end_date):
    """Lee las dos líneas (LATINO y REP DOM) del gráfico corto del PDF y las alinea con la serie diaria."""
    import pymupdf as fitz, numpy as np
    doc = fitz.open(path)
    page = None
    for p in doc:
        t = p.get_text()
        if "Riesgo- País (Spread)" in t or "Riesgo-País (Spread)" in t:
            n = [d for d in p.get_drawings() if sum(1 for it in d["items"] if it[0] == "l") > 500]
            if len(n) >= 2:
                cand = (len(n[0]["items"]), p)
                if page is None or cand[0] < page[0]: page = cand  # el gráfico más corto = ventana reciente
    if page is None: raise RuntimeError("no encontré el gráfico en el PDF")
    p = page[1]
    words = p.get_text("words")
    m = re.search(r"(\d{1,2}) de ([a-záéíóú]+),? (\d{4}) - (\d{1,2}) de ([a-záéíóú]+),? (\d{4})", p.get_text())
    start = dt.date(int(m.group(3)), MES_ES[m.group(2).lower()], int(m.group(1)))
    yt = sorted((float(w[4]), (w[1] + w[3]) / 2) for w in words if re.fullmatch(r"\d{1,2}\.\d", w[4]) and w[0] < 70)
    a, b = np.polyfit([y for _, y in yt], [v for v, _ in yt], 1)
    lines = {}
    for d in p.get_drawings():
        if sum(1 for it in d["items"] if it[0] == "l") > 500:
            col = d.get("color") or (0, 0, 0)
            name = "LAT" if col[2] > 0.6 and col[0] < 0.3 else "DOM"
            pts = [d["items"][0][1]] + [it[2] for it in d["items"] if it[0] == "l"]
            lines[name] = np.array([a * pt.y + b for pt in pts]) * 100
    P = lines["DOM"]
    ex = [(d, v) for d, v in daily if d >= start and "DOM" in v]
    K = 0
    if len(ex) > 120:
        E = np.array([v["DOM"] for _, v in ex]); i = np.arange(len(E)); best = None
        for s in np.arange(0.95, 1.10, 0.002):
            for c in range(-60, 61):
                j = np.round(s * i + c).astype(int); ok = (j >= 0) & (j < len(P))
                if ok.sum() < 0.8 * len(E): continue
                err = float(np.mean(np.abs(P[j[ok]] - E[ok])))
                if best is None or err < best[0]: best = (err, s, c)
        err, s, c = best
        K = int(round(s * (len(E) - 1) + c)) + 1
        last_known = ex[-1][0]
        log(f"gráfico PDF alineado: error medio {err:.1f} pb, índice {K}/{len(P)}")
        if err > 25: log("AVISO: alineación del gráfico con error alto")
    else:
        last_known = start - dt.timedelta(days=1)
        log("AVISO: sin solapamiento con el Excel, se reparte el gráfico por días hábiles")
    bdays = []
    d = last_known
    while d < end_date:
        d += dt.timedelta(days=1)
        if d.weekday() < 5: bdays.append(d)
    rest = len(P) - K
    idx = np.linspace(0, len(bdays) - 1, rest).round().astype(int) if rest > 0 else []
    out = {}
    for k, arr in lines.items():
        mon = collections.defaultdict(list)
        for n, ii in enumerate(idx):
            mon[bdays[ii].strftime("%Y-%m")].append(float(arr[K + n]))
        out[k] = {m: round(statistics.mean(v), 1) for m, v in mon.items()}
    return out, last_known

def build_risk(tmp, hist):
    xp, pp = os.path.join(tmp, "EMBI.xlsx"), os.path.join(tmp, "EMBI.pdf")
    base = "https://cdn.bancentral.gov.do/documents/entorno-internacional/documents/"
    open(xp, "wb").write(get(base + "EMBI.xlsx")); open(pp, "wb").write(get(base + "EMBI.pdf"))
    daily = parse_xlsx(xp)
    xl_end = daily[-1][0]
    pdate, snap = parse_pdf(pp)
    log("Excel BCRD hasta", xl_end, "| PDF del", pdate, len(snap), "países")
    # historial diario
    hist = dict(hist or {})
    hist[pdate.isoformat()] = snap
    # meses completos del Excel (el último mes puede estar parcial: se usa igual)
    mon = collections.defaultdict(lambda: collections.defaultdict(list))
    for d, vals in daily:
        for k, v in vals.items(): mon[k][d.strftime("%Y-%m")].append(v)
    series = {k: {m: round(statistics.mean(v), 1) for m, v in sorted(d.items())} for k, d in mon.items()}
    xl_last_m = xl_end.strftime("%Y-%m")
    src = {k: ["BCRD (Excel)"] for k in series}
    # gráfico del PDF para RD y Latam
    try:
        ch, _ = pdf_chart(pp, daily, pdate)
        for k in ("DOM", "LAT"):
            add = {m: v for m, v in ch[k].items() if m > xl_last_m}
            series[k].update(add)
            if add: src[k].append("BCRD (gráfico PDF)")
    except Exception as e:
        log("AVISO gráfico PDF:", e)
    # BCRP para 8 países
    bc = fetch_bcrp()
    for k, s in bc.items():
        if k == "LAT": continue
        add = {m: round(v, 1) for m, v in s.items() if m > xl_last_m and k in series}
        series[k].update(add)
        if add: src[k].append("BCRP")
    # historial diario -> promedios mensuales para meses que falten
    hm = collections.defaultdict(lambda: collections.defaultdict(list))
    for ds, vals in hist.items():
        for k, v in vals.items(): hm[k][ds[:7]].append(v)
    cur_m = pdate.strftime("%Y-%m")
    for k, d in hm.items():
        if k not in series: series[k] = {}; src[k] = []
        add = {m: round(statistics.mean(v), 1) for m, v in d.items() if m not in series[k] and m != cur_m}
        series[k].update(add)
        if add and "BCRD (dato diario)" not in src[k]: src[k].append("BCRD (dato diario)")
    risk = {}
    for k, s in series.items():
        if k not in C and k not in ("LAT", "GLB"): continue
        s = {m: v for m, v in s.items() if m < cur_m}  # el mes en curso se representa con el último dato diario
        risk[k] = {"m": dict(sorted(s.items())), "last": snap.get(k), "src": src.get(k, [])}
    return risk, pdate, hist, xl_end

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True); ap.add_argument("--hist", default=None)
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    hist = {}
    if a.hist and os.path.exists(a.hist):
        try: hist = json.load(open(a.hist))
        except Exception as e: log("historial ilegible, se empieza vacío:", e)
    tmp = tempfile.mkdtemp(prefix="economiks_")
    imf, vintage = fetch_imf()
    rem = fetch_wb()
    risk, pdate, hist, xl_end = build_risk(tmp, hist)
    # último año observado del WEO: el año anterior a la edición si es de octubre, dos antes si es de abril
    m = re.search(r"(\w+) (\d{4})", vintage or "")
    yr = int(m.group(2)) if m else dt.date.today().year
    last_real = yr - 1 if (m and m.group(1).lower().startswith(("oct", "sep"))) else yr - 2
    annual = dict(imf); annual["REM"] = rem
    meses = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"]
    vint_es = vintage
    if m:
        en = {"January":1,"February":2,"March":3,"April":4,"May":5,"June":6,"July":7,"August":8,"September":9,"October":10,"November":11,"December":12}
        if m.group(1) in en: vint_es = f"{meses[en[m.group(1)]-1].capitalize()} {m.group(2)}"
    out = {"countries": C, "reference": REF, "annual": annual, "risk": risk, "riskDate": pdate.isoformat(),
           "weo": vint_es, "lastReal": last_real, "generated": dt.date.today().isoformat(),
           "riskXlsEnd": xl_end.isoformat()}
    json.dump(out, open(os.path.join(a.out, "data.json"), "w"), ensure_ascii=False, separators=(",", ":"))
    json.dump(dict(sorted(hist.items())), open(os.path.join(a.out, "historial_riesgo.json"), "w"), ensure_ascii=False, separators=(",", ":"))
    # resumen para quien corre el script
    print(json.dumps({"riskDate": out["riskDate"], "weo": vint_es, "lastReal": last_real,
                      "DOM": risk.get("DOM", {}).get("last"), "LAT": risk.get("LAT", {}).get("last"),
                      "countries_with_risk": sum(1 for k in risk if k in C), "hist_days": len(hist)}, ensure_ascii=False))

if __name__ == "__main__":
    main()
