"""Repara el proyecto VBA del registro de placas (error 429 en SAVE_PESO_CV).

Causa: el vbaProject.bin conserva módulos de documento huérfanos
(ThisWorkbook, Hoja2..Hoja5) que quedaron de un guardado con LibreOffice,
junto a los reales (ThisWorkbook1, Hoja1, Hoja6..Hoja9). En el código,
"ThisWorkbook" se resuelve al módulo huérfano, que no tiene un libro detrás,
y Excel falla con el error 429 ("El componente ActiveX no puede crear el
objeto").

Corrección:
- Se eliminan del proyecto los módulos de documento huérfanos.
- En el código, ThisWorkbook se cambia por Application.ThisWorkbook.
- Se quita la caché compilada (p-code) para que Excel recompile el código
  fuente al abrir el libro (MS-OVBA: _VBA_PROJECT versión 0xFFFF).

Uso: python3 tools/reparar_vba.py ENTRADA.xlsm SALIDA.xlsm
"""
import io
import re
import struct
import sys
import zipfile

import olefile
from oletools.olevba import decompress_stream

HUERFANOS = {"ThisWorkbook", "Hoja2", "Hoja3", "Hoja4", "Hoja5"}
CODEPAGE = "cp1252"


# ---------------------------------------------------------------- MS-OVBA
def compress(data):
    """Compresión MS-OVBA 2.4.1 (LZ77 por bloques de 4096 bytes)."""
    out = bytearray(b"\x01")
    pos = 0
    while pos < len(data):
        start = pos
        end = min(start + 4096, len(data))
        chunk = bytearray()
        while pos < end:
            flag_pos = len(chunk)
            chunk.append(0)
            flags = 0
            for bit in range(8):
                if pos >= end:
                    break
                dpos = pos - start
                # tamaño de bits según MS-OVBA 2.4.1.3.19.1
                bitcount = max((dpos - 1).bit_length(), 4) if dpos > 0 else 4
                length_mask = 0xFFFF >> bitcount
                max_len = length_mask + 3
                best_len, best_off = 0, 0
                if dpos > 0:
                    cand = start
                    lo = start
                    for c in range(pos - 1, lo - 1, -1):
                        n = 0
                        while pos + n < end and n < max_len and data[c + n] == data[pos + n]:
                            n += 1
                        if n > best_len:
                            best_len, best_off = n, pos - c
                            if n == max_len:
                                break
                if best_len >= 3:
                    token = ((best_off - 1) << (16 - bitcount)) | (best_len - 3)
                    chunk += struct.pack("<H", token)
                    flags |= 1 << bit
                    pos += best_len
                else:
                    chunk.append(data[pos])
                    pos += 1
            chunk[flag_pos] = flags
        if len(chunk) > 4096:  # bloque sin comprimir (solo válido con 4096 bytes)
            assert end - start == 4096, "bloque final no comprimible"
            out += struct.pack("<H", 0x3000 | 4095) + bytes(data[start:end])
        else:
            out += struct.pack("<H", 0xB000 | (len(chunk) - 1)) + chunk
    return bytes(out)


def parse_dir(d):
    recs, i = [], 0
    while i < len(d):
        rid, sz = struct.unpack_from("<HI", d, i)
        if rid == 0x0009:  # PROJECTVERSION: Size=4 pero trae 6 bytes
            sz = 6
        recs.append([rid, bytes(d[i + 6:i + 6 + sz]), struct.unpack_from("<I", d, i + 2)[0]])
        i += 6 + sz
    return recs


def build_dir(recs):
    return b"".join(struct.pack("<HI", rid, size) + data for rid, data, size in recs)


# ---------------------------------------------------------------- CFB (OLE)
def write_cfb(tree):
    """tree: {"PROJECT": bytes, "VBA": {"dir": bytes, ...}}  ->  bytes (CFB v3)."""
    SEC, MINI, CUTOFF = 512, 64, 4096
    ENDC, FREE, FATSECT = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD
    entries = []  # dicts: name, type, data, children

    def add(name, node):
        e = {"name": name, "children": []}
        if isinstance(node, dict):
            e["type"] = 1
            for k in node:
                e["children"].append(add(k, node[k]))
        else:
            e["type"], e["data"] = 2, node
        entries.append(e)
        return e

    root = {"name": "Root Entry", "type": 5, "children": [add(k, v) for k, v in tree.items()]}
    entries.insert(0, root)
    for idx, e in enumerate(entries):
        e["id"] = idx

    # mini stream
    ministream, minifat = bytearray(), []
    big = []
    for e in entries:
        if e["type"] != 2:
            continue
        data = e["data"]
        if len(data) < CUTOFF:
            n = max(1, -(-len(data) // MINI)) if data else 0
            e["start"] = len(ministream) // MINI if n else ENDC
            for j in range(n):
                minifat.append(e["start"] + j + 1 if j < n - 1 else ENDC)
            ministream += data.ljust(n * MINI, b"\x00")
        else:
            big.append(e)
    dir_bytes_n = len(entries) * 128
    n_dir = -(-dir_bytes_n // SEC)
    n_minifat = -(-len(minifat) * 4 // SEC)
    n_ministream = -(-len(ministream) // SEC)
    n_big = sum(-(-len(e["data"]) // SEC) for e in big)
    n_fat = 1
    while True:
        total = n_fat + n_dir + n_minifat + n_ministream + n_big
        if total <= n_fat * (SEC // 4):
            break
        n_fat += 1
    assert n_fat <= 109
    fat = [FREE] * (n_fat * SEC // 4)
    cur = 0

    def chain(n):
        nonlocal cur
        s = cur
        for j in range(n):
            fat[cur + j] = cur + j + 1 if j < n - 1 else ENDC
        cur += n
        return s if n else ENDC

    for j in range(n_fat):
        fat[j] = FATSECT
    cur = n_fat
    dir_start = chain(n_dir)
    minifat_start = chain(n_minifat)
    ministream_start = chain(n_ministream)
    for e in big:
        e["start"] = chain(-(-len(e["data"]) // SEC))
    root["start"] = ministream_start if ministream else ENDC
    root["size"] = len(ministream)

    def key(e):
        return (len(e["name"]), e["name"].upper())

    for e in entries:  # hijos como lista enlazada por "right sibling" (todos negros)
        e.setdefault("left", FREE)
        e.setdefault("right", FREE)
        e.setdefault("child", FREE)
    for e in entries:
        kids = sorted(e["children"], key=key)
        if kids:
            e["child"] = kids[0]["id"]
            for a, b in zip(kids, kids[1:]):
                a["right"] = b["id"]

    d = bytearray()
    for e in entries:
        nm = e["name"].encode("utf-16-le") + b"\x00\x00"
        size = e.get("size", len(e.get("data", b"")))
        start = e.get("start", 0 if e["type"] == 1 else ENDC)
        if e["type"] == 1:
            start, size = 0, 0
        d += (nm.ljust(64, b"\x00") + struct.pack("<HBB", len(nm), e["type"], 1)
              + struct.pack("<III", e["left"], e["right"], e["child"]) + b"\x00" * 16
              + b"\x00" * 4 + b"\x00" * 16 + struct.pack("<IIi", start, size, 0))
    d = bytes(d).ljust(n_dir * SEC, b"\x00")
    # las entradas sin usar del último sector de directorio deben tener ids libres
    fill = bytearray(d)
    for off in range(dir_bytes_n, len(fill), 128):
        fill[off + 68:off + 80] = struct.pack("<III", FREE, FREE, FREE)
    d = bytes(fill)

    header = bytearray(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1" + b"\x00" * 16)
    header += struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6) + b"\x00" * 6
    header += struct.pack("<IIIIIIIII", 0, n_fat, dir_start, 0, CUTOFF,
                          minifat_start if n_minifat else ENDC, n_minifat, ENDC, 0)
    difat = list(range(n_fat)) + [FREE] * (109 - n_fat)
    header += struct.pack("<109I", *difat)
    assert len(header) == 512

    body = bytearray()
    body += struct.pack(f"<{len(fat)}I", *fat)
    body += d
    mf = minifat + [FREE] * (n_minifat * SEC // 4 - len(minifat))
    body += struct.pack(f"<{len(mf)}I", *mf) if mf else b""
    body += bytes(ministream).ljust(n_ministream * SEC, b"\x00")
    for e in big:
        body += e["data"].ljust(-(-len(e["data"]) // SEC) * SEC, b"\x00")
    return bytes(header + body)


# ---------------------------------------------------------------- reparación
def reparar(vba_bin):
    o = olefile.OleFileIO(vba_bin)
    recs = parse_dir(decompress_stream(bytearray(o.openstream("VBA/dir").read())))

    # separar encabezado del proyecto y módulos
    first_mod = next(i for i, r in enumerate(recs) if r[0] == 0x0019)
    head, mods, cur = recs[:first_mod], [], None
    for r in recs[first_mod:]:
        if r[0] == 0x0019:
            cur = [r]
            mods.append(cur)
        elif r[0] == 0x0010:
            tail = [r]
            break
        else:
            cur.append(r)
    keep, streams = [], {}
    for m in mods:
        name = next(r[1] for r in m if r[0] == 0x0019).decode(CODEPAGE)
        stream = next(r[1] for r in m if r[0] == 0x001A).decode(CODEPAGE)
        offset = struct.unpack("<I", next(r[1] for r in m if r[0] == 0x0031))[0]
        if name in HUERFANOS:
            continue
        raw = o.openstream("VBA/" + stream).read()
        src = decompress_stream(bytearray(raw[offset:])).decode(CODEPAGE)
        src = re.sub(r"(?<![\w.])ThisWorkbook\.", "Application.ThisWorkbook.", src)
        streams[stream] = compress(src.encode(CODEPAGE))
        for r in m:
            if r[0] == 0x0031:
                r[1] = struct.pack("<I", 0)
        keep.append(m)
    assert {"Módulo1", "Botones", "ThisWorkbook1"} <= set(streams)
    for r in head:
        if r[0] == 0x000F:
            r[1] = struct.pack("<H", len(keep))
    new_dir = build_dir(head + [r for m in keep for r in m] + tail)

    project = o.openstream("PROJECT").read().decode(CODEPAGE)
    lines = [ln for ln in project.split("\r\n")
             if not re.match(r"^(Document=|)(%s)(/&H|=)" % "|".join(HUERFANOS), ln)]
    project = "\r\n".join(lines).encode(CODEPAGE)

    wm = bytearray()
    for m in keep:
        name = next(r[1] for r in m if r[0] == 0x0019)
        wm += name + b"\x00" + name.decode(CODEPAGE).encode("utf-16-le") + b"\x00\x00"
    wm += b"\x00\x00"

    vba = {"dir": compress(new_dir), "_VBA_PROJECT": bytes.fromhex("cc61ffff000000")}
    vba.update(streams)
    return write_cfb({"PROJECT": project, "PROJECTwm": bytes(wm), "VBA": vba})


def main(src, dst):
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/vbaProject.bin":
                data = reparar(io.BytesIO(data))
            z.writestr(info, data, zipfile.ZIP_DEFLATED)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
