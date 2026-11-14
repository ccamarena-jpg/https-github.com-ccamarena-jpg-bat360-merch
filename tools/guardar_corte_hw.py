# -*- coding: utf-8 -*-
"""
Guarda un «corte» del Excel de Hardware (HARWARE ACTUALIZADO …) en un Excel histórico (HW_CORTES.xlsx).

- Lee la hoja BBDD del archivo más reciente de la carpeta HW (sincronizada con OneDrive en este PC).
- Si el contenido del archivo ya está guardado (mismo SHA-1), no hace nada: OneDrive cambia la fecha del archivo sin que
  cambie su contenido, así que se compara el contenido y no la fecha.
- Si es nuevo, agrega el corte a HW_CORTES.xlsx. La vista «Hardware» lee ese Excel desde SharePoint para comparar cortes.

Uso:   python guardar_corte_hw.py            (guarda si hay un corte nuevo)
       python guardar_corte_hw.py --forzar   (guarda aunque el contenido ya exista)
       python guardar_corte_hw.py --simular  (muestra qué haría, sin escribir)
Opciones: --origen <carpeta o archivo .xlsx>   --destino <HW_CORTES.xlsx>

IMPORTANTE: la lógica de lectura (normalización, activos, huella) replica la de bat-merch/hw-data.js. Si se cambia una, hay que
cambiar la otra; si no, la vista marcaría todos los PDV como «con cambio de hardware».
"""
import argparse, datetime, glob, hashlib, os, re, shutil, sys, tempfile, unicodedata
import openpyxl

ORIGEN = r"C:\Users\User\OneDrive - TRADITIONAL TRADE AUDIT S.A.C\Archivos de Alessandra - MASSIEL\BAT\BORRADOR\PLAN\PROYECTOS ACTUALES\HW"
DESTINO = r"C:\Users\User\OneDrive - TRADITIONAL TRADE AUDIT S.A.C\BAT\000. REPORTES\360BAT\hw-cortes\HW_CORTES.xlsx"
HOJA = "BBDD"
MAX_QTY = 10

COLS = {"id": "ID STORE", "zona": "LIMA/KC/NKC", "canal": "CANAL", "ciudad": "CIUDAD", "comCig": "COMUNICACION DE CIGARRERA", "cadena": "CADENA", "pos": "POS",
        "tvQ": "CANTIDAD PANTALLA", "cigA": "DETALLE DE CIGARRERA - A", "cigAQ": "CANTIDAD CIG - A", "cigB": "DETALLE DE CIGARRERA - B", "cigBQ": "CANTIDAD CIG - B",
        "cigC": "DETALLE DE CIGARRERA - C", "cigCQ": "CANTIDAD CIG - C", "comDisp": "COMUNICACION DISPENSER", "dispA": "DISPENSER A", "dispB": "DISPENSER B", "dispQ": "CANTIDAD DISP.",
        "gFmc": "G. FMC", "gVuse": "G. VUSE", "gVelo": "G. VELO", "gVuseQ": "CANT. GLORI VUSE", "vol": "VOLUMETRICO", "volQ": "CANTIDAD VOLUMETRICOS", "espejo": "ESPEJO"}
DIMS = ["ciudad", "cadena", "zona", "canal", "comCig", "comDisp"]
CAT_KEY = {"Cigarrera": "cig", "Dispenser": "disp", "Glorificador": "glor", "Volumétrico": "vol", "TV": "tv", "Espejo": "esp"}


def norm(s):
    if s is None:
        return ""
    if isinstance(s, float) and s.is_integer():
        s = int(s)
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().upper()


def txt(s):                      # equivale a String(v) de JS
    if s is None:
        return ""
    if isinstance(s, float) and s.is_integer():
        return str(int(s))
    return str(s)


def num(v):                      # parseFloat(String(v).replace(",", ".")) || 0
    m = re.match(r"\s*[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?", txt(v).replace(",", "."))
    return float(m.group(0)) if m else 0.0


def vacio(v):
    return norm(v) in ("", "N/A", "NA", "-", "SIN ACTIVO BAT", "SIN ACTIVO", "NO")


def fnum(n):                     # número como lo imprime JS
    return str(int(n)) if float(n).is_integer() else repr(float(n))


class Aviso:
    n = 0


def q(v):
    n = num(v)
    if n > MAX_QTY:
        Aviso.n += 1
        return 1.0
    return n


def assets_of(r):
    out = []
    for det, qty in (("cigA", "cigAQ"), ("cigB", "cigBQ"), ("cigC", "cigCQ")):
        if not vacio(r[det]):
            out.append(("Cigarrera", norm(r[det]), q(r[qty]) or 1.0))
    has_a, has_b, dq = (not vacio(r["dispA"])), (not vacio(r["dispB"])), q(r["dispQ"])
    if has_a or has_b:
        nb = 1.0 if has_b else 0.0
        na = max((dq or (nb + 1)) - nb, 1.0) if has_a else 0.0
        if has_a:
            out.append(("Dispenser", norm(r["dispA"]), na))
        if has_b:
            out.append(("Dispenser", norm(r["dispB"]), nb))
    if not vacio(r["gFmc"]):
        out.append(("Glorificador", "FMC · " + norm(r["gFmc"]), 1.0))
    if not vacio(r["gVuse"]):
        out.append(("Glorificador", "VUSE · " + norm(r["gVuse"]), q(r["gVuseQ"]) or 1.0))
    if not vacio(r["gVelo"]):
        out.append(("Glorificador", "VELO · " + norm(r["gVelo"]), 1.0))
    if not vacio(r["vol"]):
        out.append(("Volumétrico", "VOLUMÉTRICO", q(r["volQ"]) or 1.0))
    tv = q(r["tvQ"])
    if tv > 0:
        out.append(("TV", "TV", tv))
    if norm(r["espejo"]) == "SI":
        out.append(("Espejo", "ESPEJO", 1.0))
    return out


def huella(assets):
    """Misma huella que hw-data.js: djb2 sobre «cat:actxN» ordenado, en base 36."""
    s = "|".join(sorted("%s:%sx%s" % (c, a, fnum(n)) for c, a, n in assets))
    h = 5381
    b = s.encode("utf-16-le")
    for i in range(0, len(b), 2):
        h = (h * 33 + (b[i] | (b[i + 1] << 8))) & 0xFFFFFFFF
    d = "0123456789abcdefghijklmnopqrstuvwxyz"
    out = ""
    while True:
        out = d[h % 36] + out
        h //= 36
        if h == 0:
            return out


def leer_bbdd(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if HOJA not in wb.sheetnames:
        raise SystemExit("No encontré la hoja %s en %s" % (HOJA, path))
    rows = wb[HOJA].iter_rows(values_only=True)
    hdr, hi = None, 0
    buf = []
    for i, row in enumerate(rows):
        buf.append(row)
        h = [norm(c) for c in row]
        if "ID STORE" in h and "CIUDAD" in h and "CADENA" in h:
            hdr = h
            break
        if i > 8:
            break
    if hdr is None:
        raise SystemExit("No encontré los encabezados de la hoja BBDD.")
    ix = {k: (hdr.index(v) if v in hdr else -1) for k, v in COLS.items()}
    out, seen, dup = [], set(), 0
    for row in rows:
        g = lambda k: (row[ix[k]] if 0 <= ix[k] < len(row) else None)
        pid = txt(g("id")).strip()
        if not pid:
            continue
        if pid in seen:
            dup += 1
            continue
        seen.add(pid)
        r = {k: g(k) for k in COLS}
        a = assets_of(r)
        u = {"cig": 0.0, "disp": 0.0, "glor": 0.0, "vol": 0.0, "tv": 0.0, "esp": 0.0, "total": 0.0}
        for c, _, n in a:
            u[CAT_KEY[c]] += n
            u["total"] += n
        out.append({"id": pid, "pos": txt(g("pos")).strip(), "zona": norm(g("zona")) or "SIN ZONA", "canal": norm(g("canal")) or "SIN CANAL",
                    "ciudad": norm(g("ciudad")) or "SIN CIUDAD", "cadena": norm(g("cadena")) or "SIN CADENA",
                    "comCig": norm(g("comCig")) or "N/A", "comDisp": norm(g("comDisp")) or "N/A", "assets": a, "u": u, "huella": huella(a)})
    return out, dup


def por_dim(rows, key):
    m = {}
    for r in rows:
        o = m.setdefault(r[key], {"pdv": 0, "cig": 0.0, "disp": 0.0, "glor": 0.0, "vol": 0.0, "tv": 0.0, "esp": 0.0, "total": 0.0})
        o["pdv"] += 1
        for f in ("cig", "disp", "glor", "vol", "tv", "esp", "total"):
            o[f] += r["u"][f]
    return sorted(m.items(), key=lambda kv: (-kv[1]["pdv"], kv[0]))


def por_activo(rows):
    m = {}
    for r in rows:
        for c, a, n in r["assets"]:
            o = m.setdefault((c, a), [0, 0.0])
            o[0] += 1
            o[1] += n
    return sorted(m.items(), key=lambda kv: (kv[0][0], -kv[1][1]))


def sha1(path):
    h = hashlib.sha1()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def elegir_origen(origen):
    if os.path.isfile(origen):
        return origen
    c = [p for p in glob.glob(os.path.join(origen, "*.xls*")) if re.search(r"HAR?D?WARE", os.path.basename(p), re.I) and not os.path.basename(p).startswith("~$")]
    if not c:
        raise SystemExit("No encontré el Excel de Hardware en: " + origen)
    return max(c, key=os.path.getmtime)


ENC = {
    "Cortes": ["Corte", "Archivo", "PDV", "Cigarreras", "Dispensers", "Glorificadores", "Volumetricos", "TV", "Espejos", "Total activos", "Hash archivo", "Guardado el"],
    "Resumen": ["Corte", "Dimension", "Clave", "PDV", "Cigarreras", "Dispensers", "Glorificadores", "Volumetricos", "TV", "Espejos", "Total activos"],
    "Activos": ["Corte", "Categoria", "Activo", "PDV", "Unidades"],
    "PDV": ["Corte", "ID STORE", "POS", "Zona", "Canal", "Ciudad", "Cadena", "Activos", "Huella"],
}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--origen", default=ORIGEN)
    ap.add_argument("--destino", default=DESTINO)
    ap.add_argument("--forzar", action="store_true")
    ap.add_argument("--simular", action="store_true")
    a = ap.parse_args()

    src = elegir_origen(a.origen)
    tmp = os.path.join(tempfile.mkdtemp(prefix="hwcorte_"), "origen.xlsx")
    shutil.copy2(src, tmp)                                   # copia: evita choques si Excel tiene el archivo abierto
    huella_archivo = sha1(tmp)
    mtime = datetime.datetime.fromtimestamp(os.path.getmtime(src)).replace(microsecond=0)
    corte = mtime.isoformat(timespec="minutes")

    if os.path.exists(a.destino):
        wb = openpyxl.load_workbook(a.destino)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
    for nombre, enc in ENC.items():
        if nombre not in wb.sheetnames:
            wb.create_sheet(nombre).append(enc)
    hashes = {str(r[10]) for r in wb["Cortes"].iter_rows(min_row=2, values_only=True) if r[10]}
    cortes = {str(r[0]) for r in wb["Cortes"].iter_rows(min_row=2, values_only=True) if r[0]}
    if huella_archivo in hashes and not a.forzar:
        print("Sin cambios: el contenido de «%s» ya está guardado. No se agregó ningún corte." % os.path.basename(src))
        return 0
    if corte in cortes:                                     # mismo minuto, contenido distinto: se distingue con segundos
        corte = mtime.isoformat(timespec="seconds")

    rows, dup = leer_bbdd(tmp)
    t = {f: sum(r["u"][f] for r in rows) for f in ("cig", "disp", "glor", "vol", "tv", "esp", "total")}
    print("Archivo: %s\nCorte: %s | PDV: %d (duplicados ignorados: %d) | activos: %s | cantidades raras: %d" % (os.path.basename(src), corte, len(rows), dup, fnum(t["total"]), Aviso.n))
    if a.simular:
        print("(simulación: no se escribió nada)")
        return 0

    wb["Cortes"].append([corte, os.path.basename(src), len(rows)] + [t[f] for f in ("cig", "disp", "glor", "vol", "tv", "esp", "total")] + [huella_archivo, datetime.datetime.now().isoformat(timespec="seconds")])
    for d in DIMS:
        for k, o in por_dim(rows, d):
            wb["Resumen"].append([corte, d, k, o["pdv"]] + [o[f] for f in ("cig", "disp", "glor", "vol", "tv", "esp", "total")])
    for (c, ac), (pdv, un) in por_activo(rows):
        wb["Activos"].append([corte, c, ac, pdv, un])
    for r in rows:
        wb["PDV"].append([corte, r["id"], r["pos"], r["zona"], r["canal"], r["ciudad"], r["cadena"], r["u"]["total"], r["huella"]])
    os.makedirs(os.path.dirname(a.destino), exist_ok=True)
    try:
        wb.save(a.destino)
    except PermissionError:
        raise SystemExit("No pude escribir %s: está abierto en Excel. Ciérralo y vuelve a ejecutar." % a.destino)
    print("Corte guardado en", a.destino)
    return 0


if __name__ == "__main__":
    sys.exit(main())
