"""Cierre de turno (13:00 / 22:00) en SAVE_PESO_CV: SCRAP (K9) puede ser 0;
solo se exige Total (L9) > 0, con un único mensaje de advertencia.

Uso: python3 tools/ajustar_cierre_turno.py ENTRADA.xlsm SALIDA.xlsm
"""
import io
import os
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from reparar_vba import reconstruir  # noqa: E402

ANTES = '''    ' SCRAP (K9) y TOTAL (L9) obligatorios y mayores a 0 (base del % scrap y kg)
    Dim horaCierre As String
    horaCierre = Format(wsA.Range("F9").Value, "hh:mm")
    If horaCierre = "13:00" Or horaCierre = "22:00" Then
        If Val(CStr(wsA.Range("K9").Value)) <= 0 Then
            MsgBox "Complete 'SCRAP (Cant. Placas)' - cierre de turno " & horaCierre & " hs (debe ser mayor a 0)", vbExclamation
            Exit Sub
        End If
        If Val(CStr(wsA.Range("L9").Value)) <= 0 Then'''

DESPUES = '''    ' TOTAL (L9) obligatorio y mayor a 0; SCRAP (K9) puede ser 0 (base del % scrap y kg)
    Dim horaCierre As String
    horaCierre = Format(wsA.Range("F9").Value, "hh:mm")
    If horaCierre = "13:00" Or horaCierre = "22:00" Then
        If Val(CStr(wsA.Range("L9").Value)) <= 0 Then'''


def transformar(nombre, src):
    if nombre != "Módulo1":
        return src
    antes = ANTES.replace("\n", "\r\n")
    assert src.count(antes) == 1, "no se encontró el bloque de cierre de turno esperado"
    return src.replace(antes, DESPUES.replace("\n", "\r\n"))


def main(src, dst):
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/vbaProject.bin":
                data = reconstruir(io.BytesIO(data), transformar=transformar)
            z.writestr(info, data, zipfile.ZIP_DEFLATED)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
