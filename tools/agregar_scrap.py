"""Agrega la hoja SCRAP (tablas + gráficos) y la validación de cierre de turno.

Hoja SCRAP (mismo formato que GRÁFICO):
- Gráfico 1 + tabla "Scrap diario": últimas 30 fechas con registros.
- Gráfico 2 + tabla "Scrap mensual": enero a diciembre del año en curso
  (meses sin datos = 0, meses futuros vacíos).
- % Scrap = NC / (NC + C)  (nota B23 de CTRL_PESO_VISUAL).
- kg Scrap = placas NC x promedio del Peso STD de los tipos de placa
  registrados en esa fecha / 1000.
- Gráficos combinados: columnas = kg (eje Y izquierdo), línea = % (eje Y
  derecho).

Macro Módulo1.SAVE_PESO_CV: en los horarios 13:00 y 22:00 (cierre de turno)
exige SCRAP (K9) y Total (L9) mayores a 0.

Todas las hojas quedan protegidas con la contraseña "calidad".

Uso: python3 tools/agregar_scrap.py ENTRADA.xlsm SALIDA.xlsm
"""
import datetime as dt
import io
import os
import re
import sys
import zipfile
from xml.sax.saxutils import escape

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ajustar_grafico import excel_hash, fmt  # noqa: E402
from reparar_vba import reconstruir  # noqa: E402

PASSWORD = "calidad"
SH = "SCRAP"
T = "&apos;TABLA-PCPAL&apos;!"
TB, TE, TH, TQ, TR = (f"{T}${c}$2:${c}$2000" for c in "BEHQR")
TIPOS = "LISTADOS!$B$3:$B$12"
NDIAS = 30
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Septiembre", "Octubre", "Noviembre", "Diciembre"]
Q = "&quot;"

# filas de la hoja
R_TIT, R_INFO = 1, 2
R_G1 = (4, 22)            # gráfico diario: filas 4..21
R_D = 22                  # título tabla diaria; +1 encabezado; +2..+5 valores; +6 nota
R_G2 = (30, 48)           # gráfico mensual: filas 30..47
R_M = 48                  # título tabla mensual
FIRST_COL, LAST_COL = 3, 32  # C .. AF (30 fechas)


def col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def serial(d):
    return (d - dt.date(1899, 12, 30)).days


# ------------------------------------------------------------------ caches
def calcular(src, hoy):
    wb = openpyxl.load_workbook(src, data_only=True)
    tipos = [c.value for c in wb["LISTADOS"]["B"][2:12] if c.value]
    recs = []
    for r in wb["TABLA-PCPAL"].iter_rows(min_row=2, max_row=2000, values_only=True):
        if r[1] is None:
            recs.append(None)
            continue
        d = r[1].date() if isinstance(r[1], dt.datetime) else r[1]
        recs.append((d, r[4], r[7] or 0, r[16] or 0, r[17] or 0))

    def prom_std(d):
        vals = []
        for t in tipos:
            hs = [x[2] for x in recs if x and x[0] == d and x[1] == t]
            if hs:
                vals.append(sum(hs) / len(hs))
        return sum(vals) / len(vals) if vals else 0

    kg_rec = [(x[3] * prom_std(x[0]) / 1000 if x and x[3] else 0) for x in recs]
    fechas = sorted({x[0] for x in recs if x}, reverse=True)
    desc = fechas[:NDIAS]
    n = len(desc)
    ventana = list(reversed(desc))

    def agg(pred):
        s = sum(x[3] for x in recs if x and pred(x[0]))
        c = sum(x[4] for x in recs if x and pred(x[0]))
        kg = sum(k for x, k in zip(recs, kg_rec) if x and pred(x[0]))
        return s, c, (s / (s + c) if s + c else 0), kg

    diario = [agg(lambda d, f=f: d == f) for f in ventana]
    anio = hoy.year
    mensual = []
    for m in range(1, 13):
        ini = dt.date(anio, m, 1)
        fin = dt.date(anio + (m == 12), m % 12 + 1, 1)
        mensual.append(None if ini > hoy else agg(lambda d, a=ini, b=fin: a <= d < b))
    return dict(recs=recs, kg_rec=kg_rec, desc=desc, n=n, ventana=ventana, diario=diario,
                anio=anio, mensual=mensual, hoy=hoy)


# ------------------------------------------------------------------ estilos
def agregar_estilos(st):
    """Devuelve (styles.xml, dict de índices de estilo nuevos)."""
    # formatos numéricos
    nf = re.search(r'<numFmts count="(\d+)">(.*?)</numFmts>', st, re.S)
    ids = [int(x) for x in re.findall(r'numFmtId="(\d+)"', nf.group(2))]
    nid = max(ids) + 1
    fmts = {"fecha": (nid, "dd/mm"), "pct": (nid + 1, "0.0%"), "kg": (nid + 2, "0.00")}
    extra = "".join(f'<numFmt numFmtId="{i}" formatCode="{escape(c)}"/>' for i, c in fmts.values())
    st = st.replace(nf.group(0), f'<numFmts count="{int(nf.group(1)) + 3}">{nf.group(2)}{extra}</numFmts>')
    # fuente chica para la tabla de 30 columnas
    fo = re.search(r'<fonts count="(\d+)"([^>]*)>(.*?)</fonts>', st, re.S)
    nfont = int(fo.group(1))
    font = '<font><b/><sz val="10"/><color rgb="FF000000"/><name val="Arial"/><family val="2"/></font>'
    st = st.replace(fo.group(0), f'<fonts count="{nfont + 1}"{fo.group(2)}>{fo.group(3)}{font}</fonts>')
    # estilos de celda
    m = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', st, re.S)
    xfs = re.findall(r'<xf [^>]*?(?:/>|>.*?</xf>)', m.group(2), re.S)
    assert len(xfs) == int(m.group(1))

    def clon(base, numfmt=None, fontid=None):
        x = xfs[base]
        if numfmt is not None:
            x = re.sub(r'numFmtId="\d+"', f'numFmtId="{numfmt}"', x)
            if "applyNumberFormat" not in x:
                x = x.replace("<xf ", '<xf applyNumberFormat="1" ', 1)
        if fontid is not None:
            x = re.sub(r'fontId="\d+"', f'fontId="{fontid}"', x)
        return x

    nuevos = {
        "hdr_fecha": clon(122, fmts["fecha"][0]),
        "int": clon(128, 1, nfont),
        "pct": clon(128, fmts["pct"][0], nfont),
        "kg": clon(128, fmts["kg"][0], nfont),
        "int_m": clon(128, 1),
        "pct_m": clon(128, fmts["pct"][0]),
        "kg_m": clon(128, fmts["kg"][0]),
    }
    base = len(xfs)
    idx = {k: base + i for i, k in enumerate(nuevos)}
    st = st.replace(m.group(0), f'<cellXfs count="{base + len(nuevos)}">{m.group(2)}'
                                f'{"".join(nuevos.values())}</cellXfs>')
    # formato condicional: > 8 placas NC diarias (nota B23) en rojo
    dx = re.search(r'<dxfs count="(\d+)">(.*?)</dxfs>', st, re.S)
    dxf = ('<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill>'
           '<bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>')
    idx["dxf_alerta"] = int(dx.group(1))
    st = st.replace(dx.group(0), f'<dxfs count="{int(dx.group(1)) + 1}">{dx.group(2)}{dxf}</dxfs>')
    return st, idx


# ------------------------------------------------------------------ celdas
def c_txt(ref, s, text):
    return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t>{escape(text)}</t></is></c>'


def c_f(ref, s, formula, value):
    if value is None:
        return f'<c r="{ref}" s="{s}"><f>{formula}</f></c>'
    if isinstance(value, str):
        if value.startswith("#"):
            return f'<c r="{ref}" s="{s}" t="e"><f>{formula}</f><v>{value}</v></c>'
        return f'<c r="{ref}" s="{s}" t="str"><f>{formula}</f><v>{escape(value)}</v></c>'
    return f'<c r="{ref}" s="{s}"><f>{formula}</f><v>{fmt(value)}</v></c>'


def c_n(ref, s, value):
    return f'<c r="{ref}" s="{s}"><v>{fmt(value)}</v></c>'


def c_empty(ref, s):
    return f'<c r="{ref}" s="{s}"/>'


def etiqueta(d):
    return f"{d.day:02d}/{d.month:02d}/{d.year % 100:02d}"


def construir_hoja(k, S, ws_uid, prot):
    rows, heights, merges = {}, {}, []

    def put(cell_xml):
        ref = re.match(r'<c r="([A-Z]+)(\d+)"', cell_xml)
        rows.setdefault(int(ref.group(2)), []).append(cell_xml)

    def band(r, s, first, last, content_xml):
        put(content_xml)
        for cn in range(first + 1, last + 1):
            put(c_empty(f"{col(cn)}{r}", s))
        merges.append(f"{col(first)}{r}:{col(last)}{r}")

    n, anio = k["n"], k["anio"]
    last_date = k["ventana"][-1] if n else None
    first_date = k["ventana"][0] if n else None

    # --- encabezado (igual que GRÁFICO)
    band(R_TIT, 119, 2, LAST_COL, c_txt(f"B{R_TIT}", 119, "SCRAP DIARIO Y MENSUAL — % Scrap y kg"))
    heights[R_TIT] = 30.7
    info_f = (f"IF($AL$2=0,{Q}SIN REGISTROS{Q},{Q}Últimas {Q}&amp;$AL$2&amp;{Q} fechas con registros "
              f"(del {Q}&amp;$AU$2&amp;{Q} al {Q}&amp;INDEX($AU$2:$AU$31,$AL$2)&amp;{Q})  ·  "
              f"% Scrap = NC / (NC + C)  ·  Año {Q}&amp;$AL$3)")
    info_v = (f"Últimas {n} fechas con registros (del {etiqueta(first_date)} al {etiqueta(last_date)})"
              f"  ·  % Scrap = NC / (NC + C)  ·  Año {anio}") if n else "SIN REGISTROS"
    band(R_INFO, 118, 2, LAST_COL, c_f(f"B{R_INFO}", 118, info_f, info_v))
    heights[R_INFO] = 21.8
    heights[3] = 6

    # --- tabla diaria (horizontal, debajo del gráfico 1)
    r0 = R_D
    band(r0, 124, 2, LAST_COL, c_txt(f"B{r0}", 124, "Scrap diario"))
    heights[r0] = 19.6
    labels = ["Fecha", "Scrap NC (placas)", "Conformes C (placas)", "% Scrap", "kg Scrap"]
    for i, lab in enumerate(labels):
        put(c_txt(f"B{r0 + 1 + i}", 123, lab))
        heights[r0 + 1 + i] = 20.5
    for j in range(1, NDIAS + 1):
        cl = col(FIRST_COL + j - 1)
        jj = f"COLUMNS($C${r0 + 1}:{cl}${r0 + 1})"
        has = j <= n
        d = k["ventana"][j - 1] if has else None
        s_, c_, p_, kg_ = k["diario"][j - 1] if has else ("", "", "", "")
        put(c_f(f"{cl}{r0 + 1}", S["hdr_fecha"],
                f"IF({jj}&lt;=$AL$2,INDEX($AI$2:$AI$31,$AL$2+1-{jj}),{Q}{Q})", serial(d) if has else ""))
        h = f"{cl}${r0 + 1}"
        put(c_f(f"{cl}{r0 + 2}", S["int"], f"IF({h}={Q}{Q},{Q}{Q},SUMIFS({TQ},{TB},{h}))", s_))
        put(c_f(f"{cl}{r0 + 3}", S["int"], f"IF({h}={Q}{Q},{Q}{Q},SUMIFS({TR},{TB},{h}))", c_))
        put(c_f(f"{cl}{r0 + 4}", S["pct"],
                f"IF({h}={Q}{Q},{Q}{Q},IF({cl}{r0 + 2}+{cl}{r0 + 3}=0,0,{cl}{r0 + 2}/({cl}{r0 + 2}+{cl}{r0 + 3})))", p_))
        put(c_f(f"{cl}{r0 + 5}", S["kg"],
                f"IF({h}={Q}{Q},{Q}{Q},SUMIFS($AN$2:$AN$2000,$AM$2:$AM$2000,{h}))", kg_))
    nota_d = ("Nota: se muestran las últimas 30 fechas con registros (fechas sin registros no se incluyen; "
              "fechas con registros pero sin datos de scrap = 0). % Scrap = placas NC / (NC + C). "
              "kg Scrap = placas NC × promedio del Peso STD de los tipos de placa (moldes) registrados "
              "en la fecha / 1000. En rojo: más de 8 placas NC en el día (informar al supervisor).")
    band(r0 + 6, 125, 2, LAST_COL, c_txt(f"B{r0 + 6}", 125, nota_d))
    heights[r0 + 6] = 25.55

    # --- tabla mensual (horizontal, debajo del gráfico 2)
    r1 = R_M
    band(r1, 124, 2, LAST_COL, c_f(f"B{r1}", 124, f"{Q}Scrap mensual — {Q}&amp;$AL$3", f"Scrap mensual — {anio}"))
    heights[r1] = 19.6
    labels_m = ["Mes", "Scrap NC (placas)", "Conformes C (placas)", "% Scrap", "kg Scrap"]
    for i, lab in enumerate(labels_m):
        put(c_txt(f"B{r1 + 1 + i}", 123, lab))
        heights[r1 + 1 + i] = 20.5
    rng = lambda c1, a, b: f"{c1},{Q}&gt;={Q}&amp;{a},{c1},{Q}&lt;{Q}&amp;{b}"  # noqa: E731
    for mi in range(12):
        c1 = FIRST_COL + 2 * mi
        a, b = col(c1), col(c1 + 1)
        ini, fin = f"$AS${mi + 2}", f"$AT${mi + 2}"
        vals = k["mensual"][mi]
        s_, c_, p_, kg_ = vals if vals else ("", "", "", "")
        band(r1 + 1, 122, c1, c1 + 1, c_txt(f"{a}{r1 + 1}", 122, MESES[mi]))
        fut = f"{ini}&gt;TODAY()"
        band(r1 + 2, S["int_m"], c1, c1 + 1,
             c_f(f"{a}{r1 + 2}", S["int_m"], f"IF({fut},{Q}{Q},SUMIFS({TQ},{rng(TB, ini, fin)}))", s_))
        band(r1 + 3, S["int_m"], c1, c1 + 1,
             c_f(f"{a}{r1 + 3}", S["int_m"], f"IF({fut},{Q}{Q},SUMIFS({TR},{rng(TB, ini, fin)}))", c_))
        band(r1 + 4, S["pct_m"], c1, c1 + 1,
             c_f(f"{a}{r1 + 4}", S["pct_m"],
                 f"IF({fut},{Q}{Q},IF({a}{r1 + 2}+{a}{r1 + 3}=0,0,{a}{r1 + 2}/({a}{r1 + 2}+{a}{r1 + 3})))", p_))
        band(r1 + 5, S["kg_m"], c1, c1 + 1,
             c_f(f"{a}{r1 + 5}", S["kg_m"],
                 f"IF({fut},{Q}{Q},SUMIFS($AN$2:$AN$2000,{rng('$AM$2:$AM$2000', ini, fin)}))", kg_))
    # total del año (AA:AF)
    tot = [v for v in k["mensual"] if v]
    ts, tc = sum(v[0] for v in tot), sum(v[1] for v in tot)
    tk = sum(v[3] for v in tot)
    tp = ts / (ts + tc) if ts + tc else 0
    band(r1 + 1, 122, 27, 32, c_txt(f"AA{r1 + 1}", 122, "Total año"))
    band(r1 + 2, S["int_m"], 27, 32, c_f(f"AA{r1 + 2}", S["int_m"], f"SUM(C{r1 + 2}:Z{r1 + 2})", ts))
    band(r1 + 3, S["int_m"], 27, 32, c_f(f"AA{r1 + 3}", S["int_m"], f"SUM(C{r1 + 3}:Z{r1 + 3})", tc))
    band(r1 + 4, S["pct_m"], 27, 32, c_f(f"AA{r1 + 4}", S["pct_m"],
                                          f"IF(AA{r1 + 2}+AA{r1 + 3}=0,0,AA{r1 + 2}/(AA{r1 + 2}+AA{r1 + 3}))", tp))
    band(r1 + 5, S["kg_m"], 27, 32, c_f(f"AA{r1 + 5}", S["kg_m"], f"SUM(C{r1 + 5}:Z{r1 + 5})", tk))
    nota_m = ("Nota: año en curso. Meses sin datos = 0; meses futuros en blanco. % Scrap del mes = "
              "NC / (NC + C) del mes; kg = suma de los kg de scrap diarios del mes.")
    band(r1 + 6, 125, 2, LAST_COL, c_txt(f"B{r1 + 6}", 125, nota_m))
    heights[r1 + 6] = 25.55

    # --- columnas auxiliares ocultas (AH:AW)
    put(c_txt("AH1", 0, "k"))
    put(c_txt("AI1", 0, "Fechas (desc.)"))
    put(c_txt("AK2", 0, "N fechas"))
    put(c_f("AL2", 0, "COUNTIF($AI$2:$AI$31,&quot;&gt;0&quot;)", n))
    put(c_txt("AK3", 0, "Año"))
    put(c_f("AL3", 0, "YEAR(TODAY())", anio))
    put(c_txt("AM1", 0, "Fecha reg."))
    put(c_txt("AN1", 0, "kg scrap reg."))
    put(c_txt("AO1", 0, "Mes"))
    put(c_txt("AP1", 0, "Mes (eje)"))
    put(c_txt("AQ1", 0, "% gráfico"))
    put(c_txt("AR1", 0, "kg gráfico"))
    put(c_txt("AS1", 0, "Desde"))
    put(c_txt("AT1", 0, "Hasta (excl.)"))
    put(c_txt("AU1", 0, "Fecha (eje)"))
    put(c_txt("AV1", 0, "% gráfico"))
    put(c_txt("AW1", 0, "kg gráfico"))
    for j in range(1, NDIAS + 1):
        r = j + 1
        put(c_n(f"AH{r}", 0, j))
        dv = serial(k["desc"][j - 1]) if j <= n else 0
        if j == 1:
            put(c_f(f"AI{r}", 132, f"MAX({TB})", dv))
        else:
            put(c_f(f"AI{r}", 132, f"IF(AI{r - 1}=0,0,IFERROR(_xlfn.AGGREGATE(14,6,{TB}/({TB}&lt;AI{r - 1}),1),0))", dv))
        dcell = f"INDEX($C${R_D + 1}:$AF${R_D + 1},AH{r})"
        has = j <= n
        put(c_f(f"AU{r}", 0,
                f"IF(AH{r}&lt;=$AL$2,RIGHT({Q}0{Q}&amp;DAY({dcell}),2)&amp;{Q}/{Q}&amp;RIGHT({Q}0{Q}&amp;MONTH({dcell}),2)"
                f"&amp;{Q}/{Q}&amp;RIGHT(YEAR({dcell}),2),NA())",
                etiqueta(k["ventana"][j - 1]) if has else "#N/A"))
        put(c_f(f"AV{r}", 0, f"IF(AH{r}&lt;=$AL$2,INDEX($C${R_D + 4}:$AF${R_D + 4},AH{r}),NA())",
                k["diario"][j - 1][2] if has else "#N/A"))
        put(c_f(f"AW{r}", 0, f"IF(AH{r}&lt;=$AL$2,INDEX($C${R_D + 5}:$AF${R_D + 5},AH{r}),NA())",
                k["diario"][j - 1][3] if has else "#N/A"))
    for mi in range(12):
        r = mi + 2
        vals = k["mensual"][mi]
        a = col(FIRST_COL + 2 * mi)
        put(c_n(f"AO{r}", 0, mi + 1))
        put(c_txt(f"AP{r}", 0, MESES[mi]))
        put(c_f(f"AQ{r}", 0, f"IF(AS{r}&gt;TODAY(),NA(),{a}{R_M + 4})", vals[2] if vals else "#N/A"))
        put(c_f(f"AR{r}", 0, f"IF(AS{r}&gt;TODAY(),NA(),{a}{R_M + 5})", vals[3] if vals else "#N/A"))
        put(c_f(f"AS{r}", 132, f"DATE($AL$3,AO{r},1)", serial(dt.date(anio, mi + 1, 1))))
        put(c_f(f"AT{r}", 132, f"DATE($AL$3,AO{r}+1,1)",
                serial(dt.date(anio + (mi == 11), (mi + 1) % 12 + 1, 1))))
    # kg de scrap por registro (fila r de SCRAP = fila r de TABLA-PCPAL)
    for r in range(2, 2001):
        x = k["recs"][r - 2]
        fecha_f = f"IF(INDEX({T}$B:$B,ROW())={Q}{Q},{Q}{Q},INDEX({T}$B:$B,ROW()))"
        put(c_f(f"AM{r}", 132, fecha_f, serial(x[0]) if x else ""))
        cnt = f"COUNTIFS({TB},AM{r},{TE},{TIPOS})"
        qv = f"N(INDEX({T}$Q:$Q,ROW()))"
        kg_f = (f"IF(OR(AM{r}={Q}{Q},{qv}=0),0,{qv}*SUMPRODUCT(SUMIFS({TH},{TB},AM{r},{TE},{TIPOS})"
                f"/({cnt}+({cnt}=0)))/SUMPRODUCT(({cnt}&gt;0)*1)/1000)")
        put(c_f(f"AN{r}", 0, kg_f, k["kg_rec"][r - 2]))

    # --- XML de la hoja
    def colkey(cx):
        m = re.match(r'<c r="([A-Z]+)', cx).group(1)
        n_ = 0
        for ch in m:
            n_ = n_ * 26 + ord(ch) - 64
        return n_

    sheet_rows = []
    for r in sorted(set(rows) | set(heights)):
        cells = sorted(rows.get(r, []), key=colkey)
        attrs = f' ht="{heights[r]}" customHeight="1"' if r in heights else ""
        sheet_rows.append(f'<row r="{r}"{attrs}>{"".join(cells)}</row>')
    cols = ('<cols><col min="1" max="1" width="2.44140625" customWidth="1"/>'
            '<col min="2" max="2" width="20.6640625" customWidth="1"/>'
            f'<col min="{FIRST_COL}" max="{LAST_COL}" width="6.33203125" customWidth="1"/>'
            '<col min="33" max="33" width="2.44140625" customWidth="1"/>'
            '<col min="34" max="49" width="12.6640625" hidden="1" customWidth="1"/></cols>')
    cf = (f'<conditionalFormatting sqref="C{R_D + 2}:AF{R_D + 2}"><cfRule type="cellIs" '
          f'dxfId="{S["dxf_alerta"]}" priority="1" operator="greaterThan"><formula>8</formula>'
          f'</cfRule></conditionalFormatting>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" '
            'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" '
            'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" '
            'xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" '
            'xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" '
            f'xr:uid="{ws_uid}"><dimension ref="B1:AW2000"/>'
            '<sheetViews><sheetView showGridLines="0" zoomScale="90" zoomScaleNormal="90" workbookViewId="0">'
            '<selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>'
            '<sheetFormatPr baseColWidth="10" defaultRowHeight="15.05" x14ac:dyDescent="0.3"/>'
            f'{cols}<sheetData>{"".join(sheet_rows)}</sheetData>{prot}'
            f'<mergeCells count="{len(merges)}">{"".join(f"<mergeCell ref={chr(34)}{m}{chr(34)}/>" for m in merges)}'
            f'</mergeCells>{cf}'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            '<pageSetup paperSize="9" orientation="landscape"/>'
            '<drawing r:id="rId1"/></worksheet>')


# ------------------------------------------------------------------ gráficos
def _rich(text, sz, color="1F4E78", rot=None):
    body = f'<a:bodyPr rot="{rot}" vert="horz"/>' if rot is not None else "<a:bodyPr/>"
    rpr = (f'sz="{sz}" b="1"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
           '<a:latin typeface="Arial"/>')
    return (f'<c:tx><c:rich>{body}<a:lstStyle/><a:p><a:pPr><a:defRPr {rpr}</a:defRPr></a:pPr>'
            f'<a:r><a:rPr lang="es-AR" {rpr}</a:rPr><a:t>{escape(text)}</a:t></a:r></a:p></c:rich></c:tx>')


def _txpr(sz, rot=None, color="404040", bold=False):
    body = f'<a:bodyPr rot="{rot}" vert="horz"/>' if rot is not None else "<a:bodyPr/>"
    return (f'<c:txPr>{body}<a:lstStyle/><a:p><a:pPr><a:defRPr sz="{sz}" b="{int(bold)}">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="Arial"/>'
            '</a:defRPr></a:pPr><a:endParaRPr lang="es-AR"/></a:p></c:txPr>')


def _str_cache(vals):
    pts = "".join(f'<c:pt idx="{i}"><c:v>{escape(v)}</c:v></c:pt>' for i, v in enumerate(vals) if v is not None)
    return f'<c:strCache><c:ptCount val="{len(vals)}"/>{pts}</c:strCache>'


def _num_cache(vals, code):
    pts = "".join(f'<c:pt idx="{i}"><c:v>{fmt(v)}</c:v></c:pt>' for i, v in enumerate(vals) if v is not None)
    return f'<c:numCache><c:formatCode>{escape(code)}</c:formatCode><c:ptCount val="{len(vals)}"/>{pts}</c:numCache>'


def _dlbls(code, pos, color):
    return (f'<c:dLbls><c:numFmt formatCode="{escape(code)}" sourceLinked="0"/><c:spPr><a:noFill/>'
            f'<a:ln><a:noFill/></a:ln></c:spPr>{_txpr(900, color=color, bold=True)}'
            f'<c:dLblPos val="{pos}"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>'
            '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')


def chart_combo(titulo, eje_x, cat_f, cats, kg_f, kgs, pct_f, pcts, kg_name_f, pct_name_f,
                etiquetas, rot_x):
    gris = '<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></c:spPr>'
    grid = ('<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="E0E0E0"/>'
            '</a:solidFill></a:ln></c:spPr></c:majorGridlines>')
    cat = f"<c:cat><c:strRef><c:f>{cat_f}</c:f>{_str_cache(cats)}</c:strRef></c:cat>"
    bar = ('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
           f'<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>{kg_name_f}</c:f>'
           f'{_str_cache(["kg Scrap"])}</c:strRef></c:tx>'
           '<c:spPr><a:solidFill><a:srgbClr val="1F4E78"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>'
           f'{_dlbls("0.00", "outEnd", "1F4E78") if etiquetas else ""}{cat}'
           f'<c:val><c:numRef><c:f>{kg_f}</c:f>{_num_cache(kgs, "0.00")}</c:numRef></c:val></c:ser>'
           '<c:gapWidth val="60"/><c:axId val="50010001"/><c:axId val="50010002"/></c:barChart>')
    line = ('<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
            f'<c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:strRef><c:f>{pct_name_f}</c:f>'
            f'{_str_cache(["% Scrap"])}</c:strRef></c:tx>'
            '<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill>'
            '<a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr>'
            '<a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:ln w="9525"><a:solidFill>'
            '<a:srgbClr val="C00000"/></a:solidFill></a:ln></c:spPr></c:marker>'
            f'{_dlbls("0.0%", "t", "C00000") if etiquetas else ""}{cat}'
            f'<c:val><c:numRef><c:f>{pct_f}</c:f>{_num_cache(pcts, "0.0%")}</c:numRef></c:val>'
            '<c:smooth val="0"/></c:ser><c:marker val="1"/><c:axId val="50010003"/><c:axId val="50010004"/>'
            '</c:lineChart>')
    ax = (f'<c:catAx><c:axId val="50010001"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
          f'<c:delete val="0"/><c:axPos val="b"/><c:title>{_rich(eje_x, 1000)}<c:overlay val="0"/></c:title>'
          '<c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/>'
          f'<c:minorTickMark val="none"/><c:tickLblPos val="low"/>{gris}{_txpr(900, rot_x)}'
          '<c:crossAx val="50010002"/><c:crosses val="autoZero"/><c:auto val="0"/><c:lblAlgn val="ctr"/>'
          '<c:lblOffset val="100"/><c:tickLblSkip val="1"/><c:noMultiLvlLbl val="0"/></c:catAx>'
          '<c:valAx><c:axId val="50010002"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/>'
          f'</c:scaling><c:delete val="0"/><c:axPos val="l"/>{grid}'
          f'<c:title>{_rich("kg Scrap", 1000, rot=-5400000)}<c:overlay val="0"/></c:title>'
          '<c:numFmt formatCode="0.0" sourceLinked="0"/><c:majorTickMark val="none"/>'
          f'<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>{gris}{_txpr(900)}'
          '<c:crossAx val="50010001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>'
          '<c:catAx><c:axId val="50010003"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
          '<c:delete val="1"/><c:axPos val="b"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
          '<c:tickLblPos val="nextTo"/><c:crossAx val="50010004"/><c:crosses val="autoZero"/>'
          '<c:auto val="0"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
          '<c:valAx><c:axId val="50010004"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/>'
          '</c:scaling><c:delete val="0"/><c:axPos val="r"/>'
          f'<c:title>{_rich("% Scrap", 1000, color="C00000", rot=-5400000)}<c:overlay val="0"/></c:title>'
          '<c:numFmt formatCode="0.0%" sourceLinked="0"/><c:majorTickMark val="none"/>'
          f'<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>{gris}{_txpr(900, color="C00000")}'
          '<c:crossAx val="50010003"/><c:crosses val="max"/><c:crossBetween val="between"/></c:valAx>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<c:date1904 val="0"/><c:lang val="es-ES"/><c:roundedCorners val="0"/><c:chart>'
            f'<c:title>{_rich(titulo, 1400)}<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
            f'<c:plotArea><c:layout/>{bar}{line}{ax}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
            '</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/>'
            f'{_txpr(1000, color="000000")}</c:legend><c:plotVisOnly val="0"/><c:dispBlanksAs val="gap"/>'
            '</c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525">'
            '<a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>'
            '</c:chartSpace>')


def drawing(anchors):
    parts = []
    for i, (r_from, r_to, rid, name) in enumerate(anchors, start=2):
        parts.append(
            '<xdr:twoCellAnchor editAs="oneCell">'
            f'<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r_from - 1}</xdr:row>'
            '<xdr:rowOff>0</xdr:rowOff></xdr:from>'
            f'<xdr:to><xdr:col>{LAST_COL}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r_to - 1}</xdr:row>'
            '<xdr:rowOff>0</xdr:rowOff></xdr:to>'
            f'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{i}" name="{name}"/>'
            '<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/>'
            '<a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
            '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
            f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{rid}"/>'
            '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
            f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">{"".join(parts)}</xdr:wsDr>')


# ------------------------------------------------------------------ VBA
VALIDACION = '''
    ' CIERRE DE TURNO (13:00 turno TM / 22:00 turno TT):
    ' SCRAP (K9) y TOTAL (L9) obligatorios y mayores a 0 (base del % scrap y kg)
    Dim horaCierre As String
    horaCierre = Format(wsA.Range("F9").Value, "hh:mm")
    If horaCierre = "13:00" Or horaCierre = "22:00" Then
        If Val(CStr(wsA.Range("K9").Value)) <= 0 Then
            MsgBox "Complete 'SCRAP (Cant. Placas)' - cierre de turno " & horaCierre & " hs (debe ser mayor a 0)", vbExclamation
            Exit Sub
        End If
        If Val(CStr(wsA.Range("L9").Value)) <= 0 Then
            MsgBox "Complete 'Total (Cant. Placas)' - cierre de turno " & horaCierre & " hs (debe ser mayor a 0)", vbExclamation
            Exit Sub
        End If
    End If
'''


def transformar_vba(nombre, src):
    if nombre != "Módulo1":
        return src
    assert "horaCierre" not in src
    ancla = re.search(r'( *If wsA\.Range\("H13"\) = "" Then\r\n.*?End If\r\n)', src, re.S)
    assert ancla, "no se encontró la validación de H13 en SAVE_PESO_CV"
    return src[:ancla.end()] + VALIDACION.replace("\n", "\r\n") + src[ancla.end():]


# ------------------------------------------------------------------ principal
def proteger(sheet_xml, objects):
    if "<sheetProtection" in sheet_xml:
        return sheet_xml
    hv, sv, spin = excel_hash(PASSWORD)
    prot = (f'<sheetProtection algorithmName="SHA-512" hashValue="{hv}" saltValue="{sv}" '
            f'spinCount="{spin}" sheet="1" objects="{int(objects)}" scenarios="1"/>')
    return sheet_xml.replace("</sheetData>", "</sheetData>" + prot, 1)


def main(src, dst, hoy=None):
    hoy = hoy or dt.date.today()
    k = calcular(src, hoy)
    zin = zipfile.ZipFile(src)
    files = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    infos = list(zin.infolist())

    wbx = files["xl/workbook.xml"].decode()
    sheets = re.findall(r'<sheet name="([^"]+)" sheetId="\d+"(?: state="\w+")? r:id="(rId\d+)"/>', wbx)
    idx_scrap = [n for n, _ in sheets].index(SH)
    rels = files["xl/_rels/workbook.xml.rels"].decode()
    path = {n: "xl/" + re.search(r'Id="%s" [^>]*Target="([^"]+)"' % rid, rels).group(1) for n, rid in sheets}
    scrap_path = path[SH]
    old = files[scrap_path].decode()
    assert "<c " not in old, "la hoja SCRAP no está vacía"
    uid = re.search(r'xr:uid="([^"]+)"', old).group(1)

    st, S = agregar_estilos(files["xl/styles.xml"].decode())
    files["xl/styles.xml"] = st.encode()

    hv, sv, spin = excel_hash(PASSWORD)
    prot = (f'<sheetProtection algorithmName="SHA-512" hashValue="{hv}" saltValue="{sv}" '
            f'spinCount="{spin}" sheet="1" objects="0" scenarios="1"/>')
    files[scrap_path] = construir_hoja(k, S, uid, prot).encode()

    # gráficos
    n = k["n"]
    cats_d = [etiqueta(d) for d in k["ventana"]] or [None]
    kgs_d = [v[3] for v in k["diario"]] or [None]
    pcts_d = [v[2] for v in k["diario"]] or [None]
    ch_d = chart_combo(
        "Scrap diario — kg (columnas) y % (línea) · últimas 30 fechas con registros",
        "Fecha de registro", f"{SH}!Scrap_D_Fecha", cats_d, f"{SH}!Scrap_D_Kg", kgs_d,
        f"{SH}!Scrap_D_Pct", pcts_d, f"{SH}!$B${R_D + 5}", f"{SH}!$B${R_D + 4}", False, -2700000)
    ch_m = chart_combo(
        f"Scrap mensual {k['anio']} — kg (columnas) y % (línea)",
        "Mes", f"{SH}!$AP$2:$AP$13", MESES, f"{SH}!$AR$2:$AR$13",
        [v[3] if v else None for v in k["mensual"]], f"{SH}!$AQ$2:$AQ$13",
        [v[2] if v else None for v in k["mensual"]], f"{SH}!$B${R_M + 5}", f"{SH}!$B${R_M + 4}", True, None)
    charts = sorted(int(x) for x in re.findall(r"xl/charts/chart(\d+)\.xml$", "\n".join(files), re.M))
    drawings = sorted(int(x) for x in re.findall(r"xl/drawings/drawing(\d+)\.xml$", "\n".join(files), re.M))
    c1, c2 = max(charts) + 1, max(charts) + 2
    dn = max(drawings) + 1
    files[f"xl/charts/chart{c1}.xml"] = ch_d.encode()
    files[f"xl/charts/chart{c2}.xml"] = ch_m.encode()
    files[f"xl/drawings/drawing{dn}.xml"] = drawing(
        [(R_G1[0], R_G1[1], "rId1", "Grafico Scrap diario"),
         (R_G2[0], R_G2[1], "rId2", "Grafico Scrap mensual")]).encode()
    chart_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    files[f"xl/drawings/_rels/drawing{dn}.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{chart_rel}" Target="../charts/chart{c1}.xml"/>'
        f'<Relationship Id="rId2" Type="{chart_rel}" Target="../charts/chart{c2}.xml"/>'
        '</Relationships>').encode()
    sheet_rels = scrap_path.replace("worksheets/", "worksheets/_rels/") + ".rels"
    assert sheet_rels not in files
    files[sheet_rels] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
        f'Target="../drawings/drawing{dn}.xml"/></Relationships>').encode()
    ct = files["[Content_Types].xml"].decode()
    add = (f'<Override PartName="/xl/drawings/drawing{dn}.xml" '
           'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
           + "".join(f'<Override PartName="/xl/charts/chart{c}.xml" '
                     'ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
                     for c in (c1, c2)))
    files["[Content_Types].xml"] = ct.replace("</Types>", add + "</Types>").encode()

    # nombres dinámicos del gráfico diario + recálculo completo al abrir
    nombres = {"Scrap_D_Fecha": "AU", "Scrap_D_Kg": "AW", "Scrap_D_Pct": "AV"}
    dn_xml = "".join(f'<definedName name="{nm}" localSheetId="{idx_scrap}">'
                     f"OFFSET({SH}!${c}$2,0,0,MAX(1,{SH}!$AL$2),1)</definedName>"
                     for nm, c in nombres.items())
    wbx = wbx.replace('<definedName name="TM">', dn_xml + '<definedName name="TM">', 1)
    assert "Scrap_D_Fecha" in wbx
    if "fullCalcOnLoad" not in wbx:
        wbx = re.sub(r"<calcPr ", '<calcPr fullCalcOnLoad="1" ', wbx, count=1)
    files["xl/workbook.xml"] = wbx.encode()

    # protección con "calidad" en todas las hojas
    for name, p in path.items():
        if name == SH:
            continue
        files[p] = proteger(files[p].decode(), objects=name not in ("GRÁFICO",)).encode()

    # macro: validación de cierre de turno
    files["xl/vbaProject.bin"] = reconstruir(io.BytesIO(files["xl/vbaProject.bin"]),
                                             transformar=transformar_vba)

    names_out = [i.filename for i in infos]
    for extra in sorted(set(files) - set(names_out)):
        infos.append(zipfile.ZipInfo(extra, date_time=infos[0].date_time))
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for info in infos:
            z.writestr(info, files[info.filename], zipfile.ZIP_DEFLATED)
    print(f"fechas={n} ({k['ventana'][0] if n else '-'} .. {k['ventana'][-1] if n else '-'}) año={k['anio']}")
    for f, v in zip(k["ventana"], k["diario"]):
        print(f"  {f}  NC={v[0]}  C={v[1]}  %={v[2]:.2%}  kg={v[3]:.3f}")
    for m, v in zip(MESES, k["mensual"]):
        print(f"  {m:<10} " + (f"NC={v[0]} C={v[1]} %={v[2]:.2%} kg={v[3]:.3f}" if v else "(futuro)"))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
