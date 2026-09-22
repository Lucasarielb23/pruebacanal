"""Ajusta el gráfico de la hoja GRÁFICO del registro de placas visco.

- El gráfico muestra solo los últimos registros de la placa elegida
  (máximo 50; si hay menos, muestra todos los que existan).
- Series del gráfico con rangos dinámicos (nombres con DESREF/OFFSET) para
  que el eje horizontal se ajuste solo; eje vertical en escala automática.
- Hoja GRÁFICO protegida con la contraseña del libro ("calidad"), dejando
  libre la celda de selección de placa y los objetos (se puede mover y
  cambiar el tamaño del gráfico).

Uso: python3 tools/ajustar_grafico.py ENTRADA.xlsm SALIDA.xlsm
"""
import base64
import hashlib
import os
import re
import struct
import sys
import zipfile
from xml.sax.saxutils import escape

import openpyxl

MAX_REG = 50
PASSWORD = "calidad"
SHEET = "GRÁFICO"


def excel_hash(password, spin=100000):
    salt = os.urandom(16)
    h = hashlib.sha512(salt + password.encode("utf-16-le")).digest()
    for i in range(spin):
        h = hashlib.sha512(h + struct.pack("<I", i)).digest()
    return base64.b64encode(h).decode(), base64.b64encode(salt).decode(), spin


def fmt(v):
    return str(int(v)) if float(v).is_integer() else repr(float(v))


def main(src, dst):
    # ---- valores actuales (para los caches del gráfico y de las celdas) ----
    wb = openpyxl.load_workbook(src, data_only=True)
    g = wb[SHEET]
    placa = g["Q2"].value
    x6, x7, x8 = (g[c].value or 0 for c in ("X6", "X7", "X8"))
    pesos = [r[6] for r in wb["TABLA-PCPAL"].iter_rows(min_row=2, max_row=2000, values_only=True)
             if r[4] == placa]
    n = len(pesos)
    nshow = min(MAX_REG, n)
    first = n - nshow + 1
    eje_txt = f"N° de Registro (últimos {nshow} de {n} registros)"

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

    # ---- sheet1.xml (GRÁFICO) ----
    sh = files["xl/worksheets/sheet1.xml"].decode()

    def repl_row(mo):
        r = int(mo.group(1))
        body = mo.group(3)
        if r in (2, 3):  # celda de selección de placa Q2:S3 desbloqueada
            body = re.sub(r'(<c r="[QRS]%d" s=")5(")' % r, r"\g<1>%d\2" % new_s, body)
        if 2 <= r <= 301:
            body = re.sub(r'<c r="A[A-E]%d"[^>]*?(?:/>|>.*?</c>)' % r, "", body, flags=re.S)
        if 2 <= r <= MAX_REG + 1:
            k = r - 1
            rows_f = f"ROWS($AA$2:AA{r})"
            f_aa = f"IF({rows_f}&lt;=$X$12,$X$3-$X$12+{rows_f},NA())"
            f_ab = (f"IFERROR(INDEX(&apos;TABLA-PCPAL&apos;!$G$2:$G$2000,MATCH($Q$2&amp;&quot;|&quot;"
                    f"&amp;AA{r},&apos;TABLA-PCPAL&apos;!$U$2:$U$2000,0)),NA())")
            cells = []
            if k <= nshow:
                vals = [first + k - 1, pesos[first + k - 2], x6, x7, x8]
                tv = [(' t="n"', fmt(v)) for v in vals]
            else:
                tv = [(' t="e"', "#N/A")] * 5
            forms = [f_aa, f_ab] + [f"IF(ISNUMBER(AA{r}),$X${c},NA())" for c in (6, 7, 8)]
            styles = [1, 6, 3, 3, 3]
            for col, f, s, (t, v) in zip(["AA", "AB", "AC", "AD", "AE"], forms, styles, tv):
                cells.append(f'<c r="{col}{r}" s="{s}"{t}><f aca="false">{f}</f><v>{v}</v></c>')
            body += "".join(cells)
        if r == 12:
            body = ('<c r="W12" s="3" t="inlineStr"><is><t>Registros graficados (últimos, máx. 50)</t></is></c>'
                    f'<c r="X12" s="3" t="n"><f aca="false">MIN({MAX_REG},$X$3)</f><v>{nshow}</v></c>') + body
        if r == 13:
            f13 = ('IF($X$3=0,&quot;N° de Registro (sin registros)&quot;,&quot;N° de Registro (últimos &quot;'
                   '&amp;$X$12&amp;&quot; de &quot;&amp;$X$3&amp;&quot; registros)&quot;)')
            body = ('<c r="W13" s="3" t="inlineStr"><is><t>Título eje horizontal</t></is></c>'
                    f'<c r="X13" s="3" t="str"><f aca="false">{f13}</f><v>{escape(eje_txt)}</v></c>') + body
        return f'<row r="{r}"{mo.group(2)}>{body}</row>'

    sh = re.sub(r'<row r="(\d+)"([^>]*)>(.*?)</row>', repl_row, sh, flags=re.S)
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
        f"OFFSET({SHEET}!$AA$2,0,{off},MAX(1,{SHEET}!$X$12),1)</definedName>"
        for nm, off in names.items())
    wbx = wbx.replace("</definedNames>", dn + "</definedNames>", 1)
    wbx = wbx.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
    files["xl/workbook.xml"] = wbx.encode()

    # ---- chart1.xml ----
    ch = files["xl/charts/chart1.xml"].decode()

    def num_cache(vals):
        pts = "".join(f'<c:pt idx="{i}"><c:v>{fmt(v)}</c:v></c:pt>' for i, v in enumerate(vals))
        return f'<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{len(vals)}"/>{pts}</c:numCache>'

    cats = list(range(first, n + 1))
    series = {"AB": ("Graf_Peso", pesos[first - 1:]), "AC": ("Graf_LCS", [x6] * nshow),
              "AD": ("Graf_LCI", [x7] * nshow), "AE": ("Graf_Nominal", [x8] * nshow)}
    cat_xml = f"<c:cat><c:numRef><c:f>{SHEET}!Graf_N</c:f>{num_cache(cats)}</c:numRef></c:cat>"
    ch, k = re.subn(r"<c:cat>.*?</c:cat>", cat_xml, ch, flags=re.S)
    assert k == 4
    for col, (nm, vals) in series.items():
        pat = r"<c:val><c:numRef><c:f>%s!\$%s\$2:\$%s\$301</c:f>.*?</c:numRef></c:val>" % (SHEET, col, col)
        new = f"<c:val><c:numRef><c:f>{SHEET}!{nm}</c:f>{num_cache(vals)}</c:numRef></c:val>"
        ch, k = re.subn(pat, new, ch, flags=re.S)
        assert k == 1, col
    # título del eje horizontal vinculado a X13
    old_title = re.search(r"(<c:catAx>.*?<c:title><c:tx>)(<c:rich>.*?</c:rich>)(</c:tx>)", ch, re.S)
    rich = old_title.group(2)
    rpr = re.search(r"<a:defRPr.*?</a:defRPr>", rich, re.S).group(0)
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
    print(f"placa={placa} n={n} graficados={nshow} (registros {first}..{n})")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
