"""Ajusta la hoja GRÁFICO (Cp/Cpk) del registro de placas visco.

- Todos los cálculos (n, media, desviación, LCS, LCI, nominal, Cp, Cpk) y el
  gráfico usan solo los últimos registros de la placa elegida: óptimo 50
  (máximo), mínimo recomendado 25 (con menos se calcula igual y se advierte).
- Eje X del gráfico rotulado M 1 ... M n (M n = registro más reciente).
- Series del gráfico con nombres dinámicos (OFFSET) para que el eje horizontal
  se ajuste solo; eje vertical en escala automática.
- Hoja GRÁFICO protegida con la contraseña del libro ("calidad"), dejando
  libre la celda de selección de placa y los objetos (se puede mover y
  cambiar el tamaño del gráfico).

Se aplica sobre el archivo original (sin estos cambios).
Uso: python3 tools/ajustar_grafico.py ORIGINAL.xlsm SALIDA.xlsm
"""
import base64
import hashlib
import os
import re
import statistics
import struct
import sys
import zipfile
from xml.sax.saxutils import escape

import openpyxl

MAX_REG = 50
PASSWORD = "calidad"
SHEET = "GRÁFICO"
T = "&apos;TABLA-PCPAL&apos;!"
RNG = {c: f"{T}${c}$2:${c}$2000" for c in "EGHIJT"}


def excel_hash(password, spin=100000):
    salt = os.urandom(16)
    h = hashlib.sha512(salt + password.encode("utf-16-le")).digest()
    for i in range(spin):
        h = hashlib.sha512(h + struct.pack("<I", i)).digest()
    return base64.b64encode(h).decode(), base64.b64encode(salt).decode(), spin


def fmt(v):
    return str(int(v)) if float(v).is_integer() else repr(float(v))


def col_idx(ref):
    letters = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def cell(ref, s, formula=None, value=None, text=None):
    """XML de una celda: fórmula con valor cacheado o texto fijo."""
    if text is not None:
        return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t>{escape(text)}</t></is></c>'
    if isinstance(value, str):
        t, v = ("e", value) if value.startswith("#") else ("str", escape(value))
    else:
        t, v = "n", fmt(value)
    return f'<c r="{ref}" s="{s}" t="{t}"><f aca="false">{formula}</f><v>{v}</v></c>'


def set_cells(body, new_cells):
    """Reemplaza/inserta celdas en el contenido de una fila, en orden de columna."""
    cells = re.findall(r'<c r="[A-Z]+\d+"[^>]*?(?:/>|>.*?</c>)', body, re.S)
    assert "".join(cells) == body
    by_ref = {re.match(r'<c r="([A-Z]+\d+)"', c).group(1): c for c in cells}
    for c in new_cells:
        by_ref[re.match(r'<c r="([A-Z]+\d+)"', c).group(1)] = c
    return "".join(by_ref[k] for k in sorted(by_ref, key=col_idx))


def main(src, dst):
    # ---- valores actuales (para los caches de celdas y del gráfico) ----
    wb = openpyxl.load_workbook(src, data_only=True)
    placa = wb[SHEET]["Q2"].value
    regs = [r for r in wb["TABLA-PCPAL"].iter_rows(min_row=2, max_row=2000, values_only=True)
            if r[4] == placa]
    total = len(regs)
    n = min(MAX_REG, total)
    excl = total - n
    ult = regs[excl:]
    pesos = [r[6] for r in ult]
    mean = statistics.mean(pesos) if n else 0
    sd = statistics.stdev(pesos) if n >= 2 else 0
    lcs = statistics.mean(r[9] for r in ult) if n else 0
    lci = statistics.mean(r[8] for r in ult) if n else 0
    nom = statistics.mean(r[7] for r in ult) if n else 0
    cp = (lcs - lci) / (6 * sd) if sd else 0
    cpk = min((lcs - mean) / (3 * sd), (mean - lci) / (3 * sd)) if sd else 0
    if total == 0:
        msg = "SIN REGISTROS no se actualizarán los datos"
    elif total < 25:
        msg = f"Cálculo en base a {n} registros (menos de 25: resultado poco representativo)"
    elif total < 50:
        msg = f"Cálculo en base a los últimos {n} registros (mínimo 25 alcanzado; óptimo 50)"
    else:
        msg = "Cálculo en base a los últimos 50 registros"
    eje_txt = (f"Muestra: M 1 (más antiguo) … M {n} (más reciente) — últimos {n} de {total} registros"
               if total else "Muestra: sin registros")

    zin = zipfile.ZipFile(src)
    files = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    infos = zin.infolist()

    # ---- styles.xml: copia desbloqueada del estilo de la celda Q2 (s=5) ----
    st = files["xl/styles.xml"].decode()
    m = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', st, re.S)
    xfs = re.findall(r"<xf .*?</xf>", m.group(2), re.S)
    assert len(xfs) == int(m.group(1))
    unlocked = xfs[5].replace('<protection locked="true"', '<protection locked="false"')
    new_s = len(xfs)
    st = st.replace(m.group(0), f'<cellXfs count="{new_s + 1}">{m.group(2)}{unlocked}</cellXfs>')
    files["xl/styles.xml"] = st.encode()

    # ---- celdas nuevas / modificadas de GRÁFICO ----
    crit = f"$E$2:$E$2000,$Q$2"
    ifs = lambda c: (f"IF($X$3=0,0,AVERAGEIFS({RNG[c]},{RNG['E']},$Q$2,"
                     f"{RNG['T']},&quot;&gt;&quot;&amp;$X$14))")
    q = "&quot;"
    upd = {
        3: [cell("W3", 3, text="n (muestra: últimos registros, máx. 50)"),
            cell("X3", 3, f"MIN({MAX_REG},$X$12)", n)],
        4: [cell("X4", 1, ifs("G"), mean)],
        5: [cell("X5", 3, f"IF($X$3&lt;2,0,SQRT(SUMPRODUCT(({RNG['E']}=$Q$2)*({RNG['T']}&gt;$X$14)"
                          f"*({RNG['G']}-$X$4)^2)/($X$3-1)))", sd),
            cell("B5", 7, "$X$11", msg)],
        6: [cell("X6", 1, ifs("J"), lcs)],
        7: [cell("X7", 1, ifs("I"), lci)],
        8: [cell("X8", 1, ifs("H"), nom)],
        9: [cell("X9", 1, "IF($X$5=0,0,($X$6-$X$7)/(6*$X$5))", cp)],
        10: [cell("X10", 1, "IF($X$5=0,0,MIN(($X$6-$X$4)/(3*$X$5),($X$4-$X$7)/(3*$X$5)))", cpk)],
        11: [cell("X11", 1,
                  f"IF($X$12=0,{q}SIN REGISTROS no se actualizarán los datos{q},"
                  f"IF($X$12&lt;25,{q}Cálculo en base a {q}&amp;$X$3&amp;{q} registros (menos de 25: resultado poco representativo){q},"
                  f"IF($X$12&lt;50,{q}Cálculo en base a los últimos {q}&amp;$X$3&amp;{q} registros (mínimo 25 alcanzado; óptimo 50){q},"
                  f"{q}Cálculo en base a los últimos 50 registros{q})))", msg)],
        12: [cell("W12", 3, text="Total registros de la placa"),
             cell("X12", 3, f"COUNTIF({RNG['E']},$Q$2)", total)],
        13: [cell("W13", 3, text="Título eje horizontal"),
             cell("X13", 3, f"IF($X$12=0,{q}Muestra: sin registros{q},{q}Muestra: M 1 (más antiguo) … M {q}"
                            f"&amp;$X$3&amp;{q} (más reciente) — últimos {q}&amp;$X$3&amp;{q} de {q}"
                            f"&amp;$X$12&amp;{q} registros{q})", eje_txt)],
        14: [cell("W14", 3, text="Registros anteriores excluidos"),
             cell("X14", 3, "$X$12-$X$3", excl)],
        40: [cell("B40", 11, f"IF($X$3=0,{q}-{q},ROUND($X$9,3))", round(cp, 3) if n else "-"),
             cell("D40", 12, f"IF($X$3=0,{q}-{q},ROUND($X$10,3))", round(cpk, 3) if n else "-"),
             cell("F40", 12, f"IF($X$3=0,{q}-{q},ROUND($X$5,3))", round(sd, 3) if n else "-"),
             cell("H40", 12, f"IF($X$3=0,{q}-{q},ROUND($X$4,2))", round(mean, 2) if n else "-"),
             cell("J40", 12, f"IF($X$3=0,{q}-{q},ROUND($X$6,2))", round(lcs, 2) if n else "-"),
             cell("L40", 12, f"IF($X$3=0,{q}-{q},ROUND($X$7,2))", round(lci, 2) if n else "-"),
             cell("N40", 12, "$X$3", n)],
        41: [cell("B41", 13, text=(
            "Nota: todos los cálculos y el gráfico usan solo los últimos registros de la placa "
            "seleccionada (óptimo 50, mínimo 25) para evaluar la tendencia del proceso actual; "
            "M 1 es el registro más antiguo de la muestra y el de número mayor, el más reciente. "
            "LCS y LCI son el promedio de los pesos máximo y mínimo de esa muestra; el Nominal/Óptimo "
            "es el promedio del Peso STD (gr)."))],
    }
    for k in range(1, MAX_REG + 1):
        r = k + 1
        kk = f"ROWS($AA$2:AA{r})"
        if k <= n:
            vals = [f"M {k}", pesos[k - 1], lcs, lci, nom]
        else:
            vals = ["#N/A"] * 5
        forms = [f"IF({kk}&lt;=$X$3,{q}M {q}&amp;{kk},NA())",
                 f"IF({kk}&lt;=$X$3,IFERROR(INDEX({RNG['G']},MATCH($Q$2&amp;{q}|{q}&amp;($X$14+{kk}),"
                 f"{T}$U$2:$U$2000,0)),NA()),NA())"]
        forms += [f"IF({kk}&lt;=$X$3,$X${c},NA())" for c in (6, 7, 8)]
        upd.setdefault(r, []).extend(
            cell(f"{col}{r}", s, f, v)
            for col, s, f, v in zip(["AA", "AB", "AC", "AD", "AE"], [1, 6, 3, 3, 3], forms, vals))

    sh = files["xl/worksheets/sheet1.xml"].decode()

    def repl_row(mo):
        r = int(mo.group(1))
        body = mo.group(3)
        if 2 <= r <= 301:  # tabla auxiliar anterior (300 filas) -> se reemplaza por 50
            body = re.sub(r'<c r="A[A-E]%d"[^>]*?(?:/>|>.*?</c>)' % r, "", body, flags=re.S)
        if r in (2, 3):  # celda de selección de placa Q2:S3 desbloqueada
            body = re.sub(r'(<c r="[QRS]%d" s=")5(")' % r, r"\g<1>%d\2" % new_s, body)
        if r in upd:
            body = set_cells(body, upd.pop(r))
        return f'<row r="{r}"{mo.group(2)}>{body}</row>'

    sh = re.sub(r'<row r="(\d+)"([^>]*)>(.*?)</row>', repl_row, sh, flags=re.S)
    assert not upd, upd.keys()
    hv, sv, spin = excel_hash(PASSWORD)
    prot = (f'<sheetProtection algorithmName="SHA-512" hashValue="{hv}" saltValue="{sv}" '
            f'spinCount="{spin}" sheet="true" objects="false" scenarios="true"/>')
    assert "<sheetProtection" not in sh
    sh = sh.replace("</sheetData>", "</sheetData>" + prot, 1)
    files["xl/worksheets/sheet1.xml"] = sh.encode()

    # ---- workbook.xml: nombres dinámicos + recálculo al abrir ----
    wbx = files["xl/workbook.xml"].decode()
    names = {"Graf_N": 0, "Graf_Peso": 1, "Graf_LCS": 2, "Graf_LCI": 3, "Graf_Nominal": 4}
    dn = "".join(
        f'<definedName function="false" hidden="false" localSheetId="0" name="{nm}" vbProcedure="false">'
        f"OFFSET({SHEET}!$AA$2,0,{off},MAX(1,{SHEET}!$X$3),1)</definedName>"
        for nm, off in names.items())
    wbx = wbx.replace("</definedNames>", dn + "</definedNames>", 1)
    wbx = wbx.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
    files["xl/workbook.xml"] = wbx.encode()

    # ---- chart1.xml ----
    ch = files["xl/charts/chart1.xml"].decode()

    def num_cache(vals):
        pts = "".join(f'<c:pt idx="{i}"><c:v>{fmt(v)}</c:v></c:pt>' for i, v in enumerate(vals))
        return f'<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{len(vals)}"/>{pts}</c:numCache>'

    cats = "".join(f'<c:pt idx="{i}"><c:v>M {i + 1}</c:v></c:pt>' for i in range(n))
    cat_xml = (f"<c:cat><c:strRef><c:f>{SHEET}!Graf_N</c:f><c:strCache><c:ptCount val=\"{n}\"/>"
               f"{cats}</c:strCache></c:strRef></c:cat>")
    ch, k = re.subn(r"<c:cat>.*?</c:cat>", cat_xml, ch, flags=re.S)
    assert k == 4
    series = {"AB": ("Graf_Peso", pesos), "AC": ("Graf_LCS", [lcs] * n),
              "AD": ("Graf_LCI", [lci] * n), "AE": ("Graf_Nominal", [nom] * n)}
    for col, (nm, vals) in series.items():
        pat = r"<c:val><c:numRef><c:f>%s!\$%s\$2:\$%s\$301</c:f>.*?</c:numRef></c:val>" % (SHEET, col, col)
        new = f"<c:val><c:numRef><c:f>{SHEET}!{nm}</c:f>{num_cache(vals)}</c:numRef></c:val>"
        ch, k = re.subn(pat, new, ch, flags=re.S)
        assert k == 1, col
    # título del eje horizontal vinculado a X13
    old_title = re.search(r"(<c:catAx>.*?<c:title><c:tx>)(<c:rich>.*?</c:rich>)(</c:tx>)", ch, re.S)
    rpr = re.search(r"<a:defRPr.*?</a:defRPr>", old_title.group(2), re.S).group(0)
    title_ref = (f"<c:strRef><c:f>{SHEET}!$X$13</c:f><c:strCache><c:ptCount val=\"1\"/>"
                 f"<c:pt idx=\"0\"><c:v>{escape(eje_txt)}</c:v></c:pt></c:strCache></c:strRef>")
    ch = ch.replace(old_title.group(0), old_title.group(1) + title_ref + old_title.group(3), 1)
    ch = re.sub(r"(<c:catAx>.*?<c:title><c:tx>.*?</c:tx><c:overlay val=\"0\"/><c:spPr>.*?</c:spPr>)",
                lambda mo: mo.group(1) + f"<c:txPr><a:bodyPr rot=\"0\"/><a:lstStyle/><a:p><a:pPr>{rpr}</a:pPr>"
                "<a:endParaRPr lang=\"es-AR\"/></a:p></c:txPr>", ch, count=1, flags=re.S)
    # escalas automáticas (sin mín./máx. fijos) y sin graficar huecos como cero
    assert "<c:min " not in ch and "<c:max " not in ch
    ch = ch.replace('<c:dispBlanksAs val="zero"/>', '<c:dispBlanksAs val="gap"/>')
    files["xl/charts/chart1.xml"] = ch.encode()

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for info in infos:
            z.writestr(info, files[info.filename], zipfile.ZIP_DEFLATED)
    print(f"placa={placa} total={total} muestra={n} (registros {excl + 1}..{total}) "
          f"media={mean:.2f} s={sd:.3f} LCS={lcs:.2f} LCI={lci:.2f} Cp={cp:.3f} Cpk={cpk:.3f}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
