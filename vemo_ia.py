#!/usr/bin/env python3
"""
vemo_ia.py -- la IA de VEMO que redacta los boxes del MFR (P&L: Budget y MoM). Un solo archivo.
Corre 100% en tu computadora: sin Claude, sin APIs, sin internet.

Uso: doble clic en VEMO_IA.bat  ->  te pregunta empresa y mes, y deja en salidas/:
    <EMPRESA>_<AAAA-MM>_para_pegar.txt   los boxes listos para la slide
    <EMPRESA>_<AAAA-MM>.md               + "Revisar antes de pegar" + tus notas de 3 meses + lo dicho en juntas

Sin menú (para tareas automáticas):
    python vemo_ia.py comentar --empresa DAE --periodo 2026-09
    python vemo_ia.py aprender

Cómo trabaja:
  * Cifras: de la hoja "<EMPRESA> Board View (Actuals)" del FY Financials vigente (y su Budget).
  * Cada driver se desglosa hasta sumar EXACTO la variación que explica; "otros" dice de qué está hecho.
  * El porqué: tus notas de celda del mes (y del anterior para MoM), KPIs y ratios (por vehículo, por hub,
    por km, take rate, drivers, flota). Lo que no se puede sostener va a "Revisar antes de pegar".
  * Tu estilo: aprende de tus boxes publicados en los MFR (2. Monthly) y se califica contra ellos.
Nunca inventa.
"""
import copy
import datetime as _dt
import datetime as dt
import json
import math
import os
import re
import sys
import time
import unicodedata
import urllib.parse
from collections import Counter, defaultdict


# ====================================================================================================
# configuración, carpetas y periodos
# ====================================================================================================

HOME_DIR = os.environ.get("VEMO_IA_HOME") or os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HOME_DIR, "datos")
CONFIG_PATH = os.path.join(DATA_DIR, "config.json")
KNOWLEDGE_DIR = DATA_DIR
MODELS_DIR = DATA_DIR
OUTPUT_DIR_DEFAULT = os.path.join(HOME_DIR, "salidas")
JUNTAS_DIR = os.path.join(HOME_DIR, "juntas")

COMPANIES = ["DAE", "LTO", "EVF", "VCN"]

MONTH_NAMES = {
    "ene": 1, "enero": 1, "jan": 1, "january": 1, "feb": 2, "febrero": 2, "february": 2,
    "mar": 3, "marzo": 3, "march": 3, "abr": 4, "abril": 4, "apr": 4, "april": 4,
    "may": 5, "mayo": 5, "jun": 6, "junio": 6, "june": 6, "jul": 7, "julio": 7, "july": 7,
    "ago": 8, "agosto": 8, "aug": 8, "august": 8, "sep": 9, "sept": 9, "septiembre": 9,
    "setiembre": 9, "september": 9, "oct": 10, "octubre": 10, "october": 10,
    "nov": 11, "noviembre": 11, "november": 11, "dic": 12, "diciembre": 12, "dec": 12, "december": 12,
}


def ensure_dirs():
    for d in (KNOWLEDGE_DIR, MODELS_DIR, OUTPUT_DIR_DEFAULT, JUNTAS_DIR):
        os.makedirs(d, exist_ok=True)


def fix_console():
    """Windows: la consola suele ser cp1252 y se cae con acentos/símbolos."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


# --------------------------------------------------------------------------
# configuración / raíz de FP&A
# --------------------------------------------------------------------------

def load_config():
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=2)


def autodetect_root():
    """Busca la carpeta de FP&A en OneDrive sin que nadie la teclee: cualquier
    carpeta cuyo nombre termine en 'FP&A' dentro de ~/OneDrive*/ (hasta 2 niveles)."""
    home = os.path.expanduser("~")
    bases = []
    try:
        for name in os.listdir(home):
            if name.lower().startswith("onedrive"):
                bases.append(os.path.join(home, name))
    except OSError:
        pass
    for b in bases:
        try:
            for name in os.listdir(b):
                full = os.path.join(b, name)
                if os.path.isdir(full) and name.strip().upper().endswith("FP&A"):
                    return full
            for name in os.listdir(b):
                full = os.path.join(b, name)
                if os.path.isdir(full):
                    try:
                        for n2 in os.listdir(full):
                            f2 = os.path.join(full, n2)
                            if os.path.isdir(f2) and n2.strip().upper().endswith("FP&A"):
                                return f2
                    except OSError:
                        continue
        except OSError:
            continue
    return None


def resolve_root(cli_root=None):
    """Orden: --root > VEMO_ROOT > config guardada > autodetección. Guarda lo que funcione."""
    for cand in (cli_root, os.environ.get("VEMO_ROOT"), load_config().get("root"), autodetect_root()):
        if cand and os.path.isdir(cand):
            cfg = load_config()
            if cfg.get("root") != cand:
                cfg["root"] = cand
                try:
                    save_config(cfg)
                except OSError:
                    pass
            return cand
    return None


# --------------------------------------------------------------------------
# periodo
# --------------------------------------------------------------------------

def parse_periodo(txt, today=None):
    """'2026-08' | '2026/8' | '08-2026' | '8/2026' | 'agosto 2026' | 'ago-26' | 'agosto'
    -> 'YYYY-MM'. Si solo viene el mes, se asume el año más reciente que no sea futuro."""
    today = today or _dt.date.today()
    t = (txt or "").strip().lower().replace("_", "-").replace("/", "-").replace(".", "-")
    t = re.sub(r"\s+", " ", t)
    m = re.fullmatch(r"(\d{4})-(\d{1,2})", t)
    if m and 1 <= int(m.group(2)) <= 12:
        return f"{int(m.group(1))}-{int(m.group(2)):02d}"
    m = re.fullmatch(r"(\d{1,2})-(\d{4})", t)
    if m and 1 <= int(m.group(1)) <= 12:
        return f"{int(m.group(2))}-{int(m.group(1)):02d}"
    m = re.fullmatch(r"([a-záéíóú]+)[ \-]*(\d{2,4})?", t)
    if m and m.group(1) in MONTH_NAMES:
        mon = MONTH_NAMES[m.group(1)]
        if m.group(2):
            y = int(m.group(2))
            y = 2000 + y if y < 100 else y
        else:
            y = today.year if mon <= today.month else today.year - 1
        return f"{y}-{mon:02d}"
    raise ValueError(f"No entendí el periodo {txt!r}. Ejemplos válidos: 2026-08, agosto 2026, ago-26, 8/2026")


def prev_period(period):
    y, m = int(period[:4]), int(period[5:7])
    return f"{y - 1}-12" if m == 1 else f"{y}-{m - 1:02d}"


# --------------------------------------------------------------------------
# archivos candidatos por negocio
# --------------------------------------------------------------------------

_TREE_CACHE = {}
TREE_INDEX_MAX_AGE_S = 24 * 3600  # recorrer OneDrive completo tarda 1-3 min: se reutiliza un día


SCAN_EXTS = (".xlsx", ".pptx", ".xlsm", ".pdf", ".docx", ".md", ".txt")


def scan_tree(root, fresh=False):
    """Camina `root` y devuelve [(ruta, mtime, tamaño)] de todos los .xlsx y .pptx (sin carpetas de respaldo ni
    temporales ~$). En una carpeta grande de OneDrive el recorrido es lo MÁS lento (~1 min en miles de archivos),
    así que el resultado se guarda en knowledge/tree_index.json y se reutiliza por 2 horas; `fresh=True` fuerza
    recorrerlo de nuevo (lo usa `aprender` y cualquier búsqueda que no encuentre lo que pide)."""
    if root in _TREE_CACHE and not fresh:
        return _TREE_CACHE[root]
    idx_path = os.path.join(KNOWLEDGE_DIR, "tree_index.json")
    if not fresh:
        try:
            with open(idx_path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            import time as _t
            if data.get("root") == root and _t.time() - data.get("ts", 0) < TREE_INDEX_MAX_AGE_S \
                    and data.get("exts") == list(SCAN_EXTS):
                out = [(os.path.join(root, *rel.split("/")), mt, sz) for rel, mt, sz in data["files"]]
                _TREE_CACHE[root] = out
                return out
        except Exception:
            pass
    out = []
    stack = [root]
    while stack:  # scandir: en Windows la fecha/tamaño vienen con el listado (sin un stat por archivo)
        d = stack.pop()
        try:
            it = os.scandir(d)
        except OSError:
            continue
        with it:
            for e in it:
                try:
                    if e.is_dir(follow_symlinks=False):
                        if not any(x in e.name.upper() for x in ("BACKUP", "RESPALDO", "__PYCACHE__")):
                            stack.append(e.path)
                        continue
                    fl = e.name.lower()
                    if e.name.startswith("~$") or not fl.endswith(SCAN_EXTS):
                        continue
                    st = e.stat(follow_symlinks=False)
                except OSError:
                    continue
                out.append((e.path, st.st_mtime, st.st_size))
    _TREE_CACHE[root] = out
    try:
        import time as _t
        os.makedirs(KNOWLEDGE_DIR, exist_ok=True)
        with open(idx_path, "w", encoding="utf-8") as fh:
            json.dump({"root": root, "ts": _t.time(), "exts": list(SCAN_EXTS),
                       "files": [[os.path.relpath(p, root).replace("\\", "/"), mt, sz] for p, mt, sz in out]}, fh)
    except OSError:
        pass
    return out


def list_decks(root, include_regex=r"Financial Review|Board|BOARD"):
    rx = re.compile(include_regex, re.I)
    return sorted(p for p, _, _ in scan_tree(root) if p.lower().endswith(".pptx") and rx.search(p))


def candidate_files(root, company, kw_map, exclude_kw, year=None):
    """TODOS los 'FY Financials' del negocio bajo `root`, mejor candidato primero:
    primero los que traen el año pedido en la ruta o el nombre, luego el más
    reciente por fecha de modificación. Así 'comentar DAE 2025-06' no usa el
    archivo de 2026 y 'comentar DAE 2026-08' no usa uno viejo."""
    found = []
    for full, mt, _sz in scan_tree(root):
        f = os.path.basename(full)
        fu = f.upper()
        if not fu.endswith(".XLSX"):
            continue
        if "FY FINANCIALS" not in fu.replace("  ", " "):
            continue
        if any(x in fu for x in exclude_kw):
            continue
        if not any(k in fu for k in kw_map[company]):
            continue
        ymatch = 1 if (year and str(year) in full.replace("\\", "/")) else 0
        found.append((ymatch, mt, full))
    found.sort(reverse=True)
    return [f for _, _, f in found]


# ====================================================================================================
# juntas de Granola (carpeta juntas/)
# ====================================================================================================

try:
    from neural_driver_net import text_terms
except Exception:  # pragma: no cover
    def text_terms(t):
        return set(re.findall(r"[a-záéíóúñü]{3,}", (t or "").lower()))


MEET_MONTHS = {
    "enero": 1, "january": 1, "febrero": 2, "february": 2, "marzo": 3, "march": 3, "abril": 4, "april": 4,
    "mayo": 5, "may": 5, "junio": 6, "june": 6, "julio": 7, "july": 7, "agosto": 8, "august": 8,
    "septiembre": 9, "setiembre": 9, "september": 9, "octubre": 10, "october": 10, "noviembre": 11,
    "november": 11, "diciembre": 12, "december": 12,
}
MONTHS_SHORT = {"ene": 1, "jan": 1, "feb": 2, "mar": 3, "abr": 4, "apr": 4, "jun": 6, "jul": 7, "ago": 8,
                "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dic": 12, "dec": 12}
EN_MON = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7, "aug": 8, "sep": 9,
          "oct": 10, "nov": 11, "dec": 12}

# cómo aparece cada negocio en las notas (Granola transcribe de oído: "DAI", "el Tío", "BCI"...)
COMPANY_ALIASES = {
    "DAE": r"\bdae\b|\bdai\b|\bdare\b",
    "LTO": r"\blto\b|\btio\b|lease to own|fimeco|financiero de credito",
    "EVF": r"\bevf\b|\bev ?fleets?\b|\bev flits\b|\bepc\b|\bdhl\b|kia pesqueria|tesla la paz",
    "VCN": r"\bvcn\b|\bbcn\b|\bbci\b",
}
LINE_KEYWORDS = {
    "revenue": r"\brevenues?\b|\bingresos?\b|interest income|\bventas\b|\bflips\b",
    "opex": r"\bopex\b|\bcogs\b|mantenimiento|maintenance|towing|towning|grua|multa|corralon|telemetr|deducible|"
            r"recoloca|taller|parking|estacionamiento|driver payments|bonos? ",
    "gross_profit": r"gross profit|margen financiero|utilidad bruta|financial margin",
    "sga": r"sg&a|\bsga\b|g&a|nomina|payroll|salari|headcount|legal and professional|honorarios|tecnologia|"
           r"\brenta\b|\brent\b|office|marketing|liquidacion|seguro|insurance",
    "ebitda": r"ebitda",
    "net_income": r"net income|utilidad neta|perdida neta|\bnet-net\b",
}
NEXT_STEPS_RX = re.compile(r"^\s*(pr[oó]ximos pasos|next steps|siguientes? pasos?|pendientes|action items)\b", re.I)


def _mplain(s):
    s = unicodedata.normalize("NFKD", s or "")
    return "".join(c for c in s if not unicodedata.combining(c)).lower()


def _unescape(s):
    return re.sub(r"\\([&~*_#\[\]()$\-+.!])", r"\1", s or "")


# --------------------------------------------------------------------------
# cifras dichas en la junta
# --------------------------------------------------------------------------
_AMT_RX = re.compile(
    r"(?P<neg>-|−|menos\s+)?\$?\s?(?P<num>\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)\s?"
    r"(?P<unit>millones|millón|millon|million|mdp|mm|m(?![a-z²])|k(?![a-z])|mil(?![a-z]))?",
    re.I)


def amounts_ctx(text, heading=""):
    """[(monto, líneas_de_P&L_de_su_cláusula)]: cada cifra con la línea de la que habla SU cláusula (lo que va
    entre ';' o '. '). Así '2.6 millones' de 'Ingresos: 2.6 millones; margen financiero: 1 millón' cuenta
    para Revenue y no para Gross Profit. Si la cláusula no nombra línea, hereda la del encabezado."""
    head_lines = _lines_in(heading.split(" / ")[-1]) if heading else set()
    out = []
    for clause in re.split(r";|\.\s|\n", text or ""):
        ls = _lines_in(clause) or head_lines
        for a in amounts(clause):
            out.append((a, ls))
    return out


def amounts(text):
    """Montos en MXN (valor absoluto) mencionados en el texto: '5 millones', '1.5M', '700K', '700,000 pesos',
    '-15.1M'. Ignora años, porcentajes y números chicos sin unidad (horas, unidades, precios unitarios)."""
    out = []
    for m in _AMT_RX.finditer(text or ""):
        num, unit = m.group("num"), (m.group("unit") or "").lower()
        tail = text[m.end():m.end() + 2]
        if tail.startswith("%"):
            continue
        try:
            v = float(num.replace(",", ""))
        except ValueError:
            continue
        if unit in ("millones", "millón", "millon", "million", "mdp", "mm", "m"):
            v *= 1e6
        elif unit in ("k", "mil"):
            v *= 1e3
        elif "," not in num:
            continue  # número suelto sin unidad (318 autos, 2026, 1.5 hrs): no es un monto
        if v < 50_000 or (1900 <= v <= 2100):
            continue
        out.append(v)
    return out


def _close_amount(a, target):
    """Coincidencia de monto: ±5% (mín. 50k) y solo para movimientos de >= 1m, donde una coincidencia casual
    es improbable. (Las juntas redondean: '~5 millones' vs 4.9m del Excel.)"""
    t = abs(target)
    return t >= 1_000_000 and abs(a - t) <= max(0.05 * t, 50_000)


# --------------------------------------------------------------------------
# lectura de archivos de juntas
# --------------------------------------------------------------------------

def _parse_date(s):
    s = (s or "").strip()
    m = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.search(r"([A-Za-z]{3})[a-z]*\.? (\d{1,2}), (\d{4})", s)  # "Sep 30, 2026 3:00 PM"
    if m and m.group(1).lower() in EN_MON:
        return dt.date(int(m.group(3)), EN_MON[m.group(1).lower()], int(m.group(2)))
    m = re.search(r"(\d{1,2})[/ -](\d{1,2})[/ -](\d{4})", s)
    if m:
        try:
            return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    return None


def _juntas_dirs():
    dirs = [JUNTAS_DIR]
    for d in (load_config().get("juntas_dirs") or []):
        if d and os.path.isdir(d) and d not in dirs:
            dirs.append(d)
    return dirs


def list_meeting_files():
    out = []
    for d in _juntas_dirs():
        if not os.path.isdir(d):
            continue
        for dp, _dn, fn in os.walk(d):
            for f in fn:
                if f.lower().endswith((".md", ".txt", ".json")) and not f.startswith(("~$", ".")) and f.lower() != "leeme.md":
                    out.append(os.path.join(dp, f))
    return sorted(out)


def _read_text_meeting(path):
    with open(path, "r", encoding="utf-8", errors="replace") as fh:
        raw = fh.read()
    meta, body = {}, raw
    m = re.match(r"\s*---\s*\n(.*?)\n---\s*\n(.*)", raw, re.S)
    if m:
        for line in m.group(1).splitlines():
            if ":" in line:
                k, v = line.split(":", 1)
                meta[_mplain(k.strip())] = v.strip()
        body = m.group(2)
    title = meta.get("titulo") or meta.get("title")
    if not title:
        first = next((l.strip("# ").strip() for l in body.splitlines() if l.strip()), "")
        title = first[:120] or os.path.splitext(os.path.basename(path))[0]
    date = _parse_date(meta.get("fecha") or meta.get("date") or "") or _parse_date(os.path.basename(path))
    if not date:
        date = dt.date.fromtimestamp(os.path.getmtime(path))
    return [{"title": title, "date": date, "url": meta.get("url", ""), "participants": meta.get("participantes", ""),
             "body": body, "file": path}]


def _read_json_meeting(path):
    with open(path, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    items = data if isinstance(data, list) else [data]
    out = []
    for it in items:
        body = it.get("summary") or it.get("notes") or it.get("body") or ""
        if not body.strip():
            continue
        out.append({"title": it.get("title", os.path.basename(path)), "date": _parse_date(it.get("date", "")) or
                    dt.date.fromtimestamp(os.path.getmtime(path)), "url": it.get("url", ""),
                    "participants": ", ".join(it.get("participants") or []) if isinstance(it.get("participants"), list)
                    else (it.get("participants") or ""), "body": body, "file": path})
    return out


def load_meetings():
    meetings, seen = [], set()
    for p in list_meeting_files():
        try:
            items = _read_json_meeting(p) if p.lower().endswith(".json") else _read_text_meeting(p)
        except Exception:
            continue
        for mt in items:
            key = (mt["title"].strip().lower(), mt["date"].isoformat())
            if key in seen:
                continue
            seen.add(key)
            meetings.append(mt)
    meetings.sort(key=lambda m: m["date"])
    return meetings


# --------------------------------------------------------------------------
# fragmentos etiquetados
# --------------------------------------------------------------------------

def _companies_in(text):
    t = _mplain(text)
    return {c for c, rx in COMPANY_ALIASES.items() if re.search(rx, t)}


def _lines_in(text):
    t = _mplain(text)
    return {l for l, rx in LINE_KEYWORDS.items() if re.search(rx, t)}


_CMP_RX = re.compile(r"\b(vs\.?|versus|contra|than|compared to|frente a|respecto a|desde|from)\s+(el mes de\s+|(?:el\s+)?cierre de\s+)?([a-z]+)")


def _periods_in(text, mdate):
    # "vs julio" / "contra junio" nombran el mes de COMPARACIÓN, no del que se habla
    t = _CMP_RX.sub(lambda m: m.group(1) if (m.group(3) in MEET_MONTHS) else m.group(0), _mplain(text))
    out = set()
    for name, mon in list(MEET_MONTHS.items()):
        if re.search(rf"\b{name}\b", t):
            if name == "may" and not re.search(r"\bmay\b(?! be)", t):
                continue
            y = mdate.year if mon <= mdate.month else mdate.year - 1
            out.add(f"{y}-{mon:02d}")
    for m in re.finditer(r"\b(20\d{2})[-/](\d{1,2})\b", t):
        if 1 <= int(m.group(2)) <= 12:
            out.add(f"{m.group(1)}-{int(m.group(2)):02d}")
    return out


def _default_period(mdate):
    """Las juntas de cierre ocurren el mes siguiente: una junta del 17-sep habla, salvo que diga otra cosa, de agosto."""
    y, m = mdate.year, mdate.month
    return f"{y - 1}-12" if m == 1 else f"{y}-{m - 1:02d}"


def chunk_meeting(mt):
    title_cos = _companies_in(mt["title"])
    title_periods = _periods_in(mt["title"], mt["date"])
    chunks, heads = [], []
    cur = None

    def flush():
        nonlocal cur
        if cur and cur["text"].strip():
            chunks.append(cur)
        cur = None

    for raw in _unescape(mt["body"]).splitlines():
        line = raw.rstrip()
        if not line.strip():
            continue
        hm = re.match(r"^(#{1,4})\s+(.*)", line.strip())
        if hm:
            flush()
            lvl = len(hm.group(1))
            heads = [h for h in heads if h[0] < lvl] + [(lvl, hm.group(2).strip())]
            continue
        bm = re.match(r"^(\s*)[-*•]\s+(.*)", line)
        indent = len(bm.group(1)) if bm else None
        txt = (bm.group(2) if bm else line.strip()).strip()
        if bm and indent == 0 or (not bm and not raw.startswith(" ") and cur is None):
            flush()
            cur = {"text": txt, "heading": " / ".join(h for _l, h in heads)}
        elif cur is not None:
            cur["text"] += ("; " if bm else " ") + txt
        else:
            cur = {"text": txt, "heading": " / ".join(h for _l, h in heads)}
    flush()

    out = []
    for i, c in enumerate(chunks):
        text = c["text"].replace("**", "").strip()
        head = c["heading"]
        cos = _companies_in(text) or _companies_in(head) or title_cos
        per = _periods_in(text, mt["date"]) or _periods_in(head, mt["date"]) or title_periods
        is_next = bool(NEXT_STEPS_RX.search(head.split(" / ")[-1])) if head else False
        out.append({
            "id": f"{mt['date'].isoformat()}#{i}", "text": text, "heading": head, "meeting": mt["title"],
            "date": mt["date"].isoformat(), "url": mt.get("url", ""), "file": mt.get("file", ""),
            "companies": sorted(cos), "lines": sorted(_lines_in(text + " " + head.split(" / ")[-1])),
            "periods": sorted(per), "default_period": _default_period(mt["date"]),
            "amounts": amounts(text), "amounts_ctx": [(a, sorted(ls)) for a, ls in amounts_ctx(text, head)],
            "kind": "pendiente" if is_next else "hecho",
        })
    return out


class MeetingMemory:
    def __init__(self, chunks, n_meetings=0, date_range=None):
        self.chunks = chunks
        self.n_meetings = n_meetings
        self.date_range = date_range
        df = Counter()
        self._tf = []
        for c in chunks:
            terms = set(text_terms(c["text"] + " " + c["heading"].split(" / ")[-1]))
            self._tf.append(terms)
            df.update(terms)
        n = max(len(chunks), 1)
        self.idf = {t: math.log((1 + n) / (1 + k)) + 1.0 for t, k in df.items()}
        self._norm = [math.sqrt(sum(self.idf[t] ** 2 for t in tf)) or 1.0 for tf in self._tf]

    @classmethod
    def load(cls):
        ms = load_meetings()
        chunks = [c for mt in ms for c in chunk_meeting(mt)]
        rng = (ms[0]["date"].isoformat(), ms[-1]["date"].isoformat()) if ms else None
        return cls(chunks, len(ms), rng)

    def __len__(self):
        return len(self.chunks)

    def _sim(self, i, query_terms, qn):
        tf = self._tf[i]
        dot = sum(self.idf.get(t, 1.0) ** 2 for t in query_terms if t in tf)
        return dot / (qn * self._norm[i]) if dot else 0.0

    def _about(self, c, company, period):
        if company and company not in c["companies"]:
            return False
        if period and not (period in c["periods"] or (not c["periods"] and c["default_period"] == period)):
            return False
        return True

    def for_line(self, company, period, canon, deltas=(), query="", k=2):
        """Fragmentos de juntas de `company` sobre `period` relevantes a la línea `canon`. Cada uno trae
        `match`: las cifras de la junta que coinciden con alguno de `deltas` (lista de (etiqueta, valor_mxn))."""
        q = set(text_terms(query)) if query else set()
        qn = math.sqrt(sum(self.idf.get(t, 1.0) ** 2 for t in q)) or 1.0
        scored = []
        for i, c in enumerate(self.chunks):
            if c["kind"] != "hecho" or not self._about(c, company, period):
                continue
            match = []
            for lab, val in deltas:
                if val is None:
                    continue
                for a, als in c["amounts_ctx"]:
                    if canon in als and _close_amount(a, val):
                        match.append((lab, a, val))
                        break
            on_line = canon in c["lines"]
            if not on_line and not (match and not c["lines"]):
                continue
            score = 1.0 + (2.0 if match else 0.0) + (0.5 if period in c["periods"] else 0.0) + (self._sim(i, q, qn) if q else 0.0)
            scored.append((score, i, match))
        scored.sort(key=lambda x: -x[0])
        return [dict(self.chunks[i], match=m, score=round(s, 2)) for s, i, m in scored[:k]]

    def pendientes(self, company, period, k=6):
        out = [c for c in self.chunks if c["kind"] == "pendiente" and self._about(c, company, period)]
        return out[:k]

    def search(self, query, company=None, period=None, k=8):
        q = set(text_terms(query or ""))
        qn = math.sqrt(sum(self.idf.get(t, 1.0) ** 2 for t in q)) or 1.0
        scored = []
        for i, c in enumerate(self.chunks):
            if not self._about(c, company, period):
                continue
            s = self._sim(i, q, qn) if q else 0.01
            if s > 0:
                scored.append((s, c["date"], i))
        if q:
            scored.sort(key=lambda x: (-x[0], x[1]))
        else:  # sin tema: lo más reciente primero
            scored.sort(key=lambda x: x[1], reverse=True)
        return [dict(self.chunks[i], similarity=round(s, 2)) for s, _d, i in scored[:k]]

    def records_for_learning(self):
        """Fragmentos como 'records' {text} para el confirmador numérico de vemo_learn."""
        return [{"text": c["text"], "companies": c["companies"]} for c in self.chunks if c["kind"] == "hecho"]


def fmt_m(v):
    return f"{v / 1e6:.1f}m" if v >= 1e6 else f"{v / 1e3:.0f}k"


def write_meeting_file(meeting, folder=JUNTAS_DIR):
    """Guarda una junta (dict con title/date/url/participants/summary) como .md con encabezado."""
    os.makedirs(folder, exist_ok=True)
    d = meeting["date"] if isinstance(meeting["date"], str) else meeting["date"].isoformat()
    safe = re.sub(r"[^\w\- ]+", "", _mplain(meeting["title"])).strip().replace(" ", "_")[:60] or "junta"
    path = os.path.join(folder, f"{d}_{safe}.md")
    parts = meeting.get("participants") or ""
    if isinstance(parts, list):
        parts = ", ".join(parts)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(f"---\ntitulo: {meeting['title']}\nfecha: {d}\nurl: {meeting.get('url', '')}\n"
                 f"participantes: {parts}\nfuente: granola\n---\n{meeting['summary'].strip()}\n")
    return path


# ====================================================================================================
# notas de celda -> el "por qué" en inglés
# ====================================================================================================

MONTHS_EN = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
             "October", "November", "December"]

# conceptos -> inglés (regex sobre el texto del renglón de la nota, sin acentos, minúsculas)
GLOSSARY = [
    (r"cobranding|co-branding", "the Uber co-branding"),
    (r"publicidad digital", "digital advertising"),
    (r"correctivo mayor", "major corrective maintenance"),
    (r"correctivos?\b", "corrective maintenance"),
    (r"preventivos?\b", "preventive maintenance"),
    (r"honorarios", "professional fees"),
    (r"\bemc\b|electromecanica", "Vemo Electromecánica (EMC)"),
    (r"reembolsos? (de )?(multas )?gn\b|reembolsos? .*guardia nacional", "National Guard fine reimbursements"),
    (r"multas? (de )?gn\b|multas? .*guardia nacional", "National Guard fines"),
    (r"reembolsos? (de )?uber", "Uber reimbursements"),
    (r"reembolsos? (de )?seguros?", "insurance reimbursements"),
    (r"reembolsos? (a|al) conductor", "driver reimbursements"),
    (r"provision.*(energia|luz|cfe)", "an energy cost provision"),
    (r"cancelacion (de (la )?)?provision", "a provision reversal"),
    (r"facturas? pendientes", "pending invoices"),
    (r"gestoria", "vehicle registration fees"),
    (r"tenencias?", "vehicle ownership taxes"),
    (r"cuotas? complementarias(?!.*\biva\b)", "supplementary IMSS contributions"),
    (r"impuesto sobre nomina", "payroll tax"),
    (r"polizas?", "insurance policies"),
    (r"deducibles?", "deductibles"),
    (r"instalacion de camara dual|camara dual", "dual-camera installation"),
    (r"recuperacion por siniestros|reembolso.*siniestro|siniestros? recuperad", "claim recoveries"),
    (r"recargos.*iva|iva.*recargos", "VAT surcharges"),
    (r"sindicato", "union fees"),
    (r"prorroga imss|imss", "IMSS advisory fees"),
    (r"implementacion.*dynamics|dynamics.*implementacion", "Dynamics implementation"),
    (r"honorarios contables|auditoria", "accounting/audit fees"),
    (r"viniles?|vinil|branding|abracam", "vinyl & branding"),
    (r"facebook|google|indeed|meta ads", "lead generation"),
    (r"corrective", "corrective maintenance"),
    (r"preventive", "preventive maintenance"),
    (r"siniestros", "accident repairs"),
    (r"telemetria", "telematics"),
    (r"resguardo", "fleet storage"),
    (r"corralon", "impound fees"),
    (r"grua", "towing"),
    (r"despensa", "groceries"),
    (r"papeleria", "stationery"),
    (r"limpieza", "cleaning"),
    (r"multa", "fines"),
    (r"baja de placas", "plate cancellations (pre-auction)"),
    (r"ubers? de empleados", "employee rides"),
    (r"comisiones bancarias", "bank fees"),
]

# conceptos que reducen el gasto (una recuperación mayor = gasto menor)
RECOVERY_RX = r"recuperacion|reembolso"
TOTAL_RX = r"^(total|neto|net|insurance neto|towing service neto|total maintenance|% de|promedio|avg|maintenance per km|third party)"


def _plain(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def note_load_glossary(home=None):
    g = list(GLOSSARY)
    try:
        p = os.path.join(home or KNOWLEDGE_DIR, "glosario_notas.json")
        if os.path.exists(p):
            with open(p, encoding="utf-8") as fh:
                extra = json.load(fh)
            g = [(k, v) for k, v in extra.items() if not k.startswith("_")] + g
    except Exception:
        pass
    try:  # términos de VEMO: solo llenan huecos (van después del glosario base)
        p = os.path.join(home or KNOWLEDGE_DIR, "glosario_vemo.json")
        if os.path.exists(p):
            with open(p, encoding="utf-8") as fh:
                g += [(k, v) for k, v in json.load(fh).items() if not k.startswith("_")]
    except Exception:
        pass
    return g


_GLOSS = None


def _gloss():
    global _GLOSS
    if _GLOSS is None:
        _GLOSS = note_load_glossary()
    return _GLOSS


GENERIC = r"^(prov\.?|pago|otros|others|varios|total)$"


def _hit(text):
    k = _plain(text)
    for rx, en in _gloss():
        if re.search(rx, k):
            return en
    return None


def _clean_name(name):
    name = re.sub(r"^\W+|[\s:;,.\-]+$", "", name.strip())
    if name.count("(") > name.count(")"):
        name = name.split("(")[0].strip()
    name = re.sub(r"\s*\(([^)]*)\)", r" \1", name).strip()
    return name


# palabras comunes en español: si el concepto de la nota no está en el glosario y no es un nombre propio
# (proveedor, plataforma), no se cita en inglés (queda en el contexto del reporte)
SPANISH_COMMON = set("""aumento disminucion incremento decremento pago pagos ajuste ajustes provision provisiones cancelacion
facturas factura pendientes pendiente reembolso reembolsos mxn mes meses total otros otras varios costo costos gasto gastos
servicio servicios renta rentas recargos actualizacion multa multas cuotas comision comisiones cargo cargos registro
reclasificacion devolucion diferencia importe monto saldo partida cuenta pagado pagada ingreso ingresos de del la el los
las por para con en y a al se que su sus como sin contrarrestado contrarrestada plataforma baja reembolo nota notas
mantenimiento compra compras tarjeta tarjetas amex cuenta""".split()) | set(MONTHS_PLAIN := ["enero", "febrero", "marzo", "abril",
    "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", "ene", "feb", "mar", "abr",
    "jun", "jul", "ago", "sep", "oct", "nov", "dic"])


EN_CONT = {"rewards", "interco", "media", "fleet", "fleets", "services", "license", "licenses", "platform", "app"}
ABBR_REJECT = {"mm", "prov", "mxn", "usd", "na", "nd", "rc", "vf"}


def _proper(name):
    """'Sofia Salud seguros empleados' -> 'Sofia Salud'; 'Facturas RC' / 'abril' / 'MXN' -> None."""
    toks = re.findall(r"[A-Za-zÁÉÍÓÚáéíóúÑñ&][\wÁÉÍÓÚáéíóúÑñ&.\-]*", name or "")
    out = []
    for t in toks:
        if _plain(t) in SPANISH_COMMON:
            if not out:
                continue  # "Pago a Proactive Risk" -> "Proactive Risk"
            break
        if not (t[:1].isupper() or (out and _plain(t) in EN_CONT)):
            break
        out.append(t)
    if not out or len(out) > 3 or _plain(out[0]) in SPANISH_COMMON or out[0].isdigit() \
            or all(_plain(t).strip(".") in ABBR_REJECT for t in out):
        return None
    return " ".join(out)


def note_english(name, seg=None):
    if seg and re.search(r"(recuperacion|reembolso)\w*\s+(de\s+\$?[\d.,]+\w*\s+)?por siniestro", _plain(seg)):
        return "claim recoveries"
    h = _hit(name)
    if h:
        return h
    name = _clean_name(name)
    hs = _hit(seg) if seg else None
    prop = _proper(name)
    if hs:
        return hs if (re.search(GENERIC, _plain(name)) or not prop) else f"{prop} {hs}"
    return prop


# --------------------------------------------------------------------------
# parseo de la nota
# --------------------------------------------------------------------------

AMT_RX = re.compile(r"(?P<sign>[+\-‑−])?\s*\$\s*(?P<num>\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)\s*(?P<suf>mm|m|k|K|M)?\b"
                    r"|(?P<sign2>[+\-‑−])?(?P<num2>\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+\.\d+|\d+)\s*(?P<suf2>mm|M|K|k|mxn)\b")


def _amount(m):
    num = m.group("num") or m.group("num2")
    suf = (m.group("suf") or m.group("suf2") or "").lower()
    sign = m.group("sign") or m.group("sign2") or ""
    v = float(num.replace(",", ""))
    if suf in ("k",):
        v *= 1e3
    elif suf in ("m", "mm"):
        v *= 1e6
    return (-v if sign in ("-", "‑", "−") else v), bool(sign)


def note_segments(note):
    if not note:
        return []
    parts = re.split(r"\s*\|\s*|\s+•\s*|^\s*•\s*|\s+-\s+(?=[A-ZÁÉÍÓÚa-z])", note)
    return [p.strip(" -•;.") for p in parts if p and p.strip(" -•;.")]


HUB_SEG = re.compile(r"^[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ .]{3,}:")


def note_items(note, label=None):
    """[(nombre, monto_abs, texto_local)] de 'Concepto: $monto' (sin totales, note_hubs, promedios ni el propio renglón)."""
    out = []
    lab = _plain(label or "")
    for seg in note_segments(note):
        if re.search(r"usd\b", seg, re.I) or HUB_SEG.match(seg):
            continue
        head = _plain(seg.split(":")[0])
        if lab and head.startswith(lab[:12]):  # renglón resumen de la propia línea ("EV Fuel Cost: -2.7M, ...")
            continue
        prev_end = 0
        ms = list(AMT_RX.finditer(seg))
        for i, m in enumerate(ms):
            raw = seg[prev_end:m.start()]
            raw = re.split(r"[(;,]|\ben\b", raw)[-1] if re.search(r"[(;,]", raw) and i > 0 else raw
            name = raw.strip(" :-•(")
            nxt = ms[i + 1].start() if i + 1 < len(ms) else len(seg)
            local = seg[prev_end:min(nxt, m.end() + 90)]
            prev_end = m.end()
            if re.search(r"^\s*(%|por unidad|por servicio|u\b|unidades|vins)", seg[m.end():m.end() + 14], re.I):
                continue
            if re.search(r"\d,\d{2}(?!\d)", seg[m.start():m.end() + 1]):  # "$115,23": monto mal capturado
                continue
            if not name:  # "Reembolso de $328k por siniestros" -> el concepto va después
                name = seg[m.end():].split(":")[0].strip(" :-•")[:60]
            pn = _plain(name)
            if not name or len(name) > 90 or re.search(TOTAL_RX, pn) or (lab and pn.startswith(lab[:12])) \
                    or re.search(r"%|\bvs\.?\b|\bmom\b", pn) or not re.search(r"[a-z]{3,}", pn):
                continue
            name = re.sub(r"^(pago de|pago a)\s+", "", name, flags=re.I)
            v, _signed = _amount(m)
            if abs(v) < 1000:
                continue
            out.append((name, abs(v), local))
    return out


def note_variations(note):
    """'MAIN COST VARIATIONS ... • Concepto: +$238,097' -> [(nombre, +cambio de costo)]."""
    if not note or not re.search(r"variation", note, re.I):
        return []
    txt = note[re.search(r"variation", note, re.I).end():]
    txt = re.split(r"TOP \d|COSTS", txt)[0]
    out = []
    for seg in note_segments(txt):
        m = re.match(r"(?P<name>[A-Za-zÁÉÍÓÚáéíóúñ][\w .áéíóúñ]+?)\s*:\s*(?P<sign>[+\-])\$(?P<num>[\d,]+)", seg)
        if m:
            v = float(m.group("num").replace(",", ""))
            out.append((m.group("name").strip(), v if m.group("sign") == "+" else -v))
    return out


def note_hubs(note):
    """'TORRE MEXICANA: -0.6M (-0.2M; ...' -> [(hub, delta en signo P&L)]."""
    out = []
    for m in re.finditer(r"([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ .]{3,}):\s*[+\-]?[\d.]+\s*([MK])\s*\(([+\-]?[\d.]+)\s*([MK])", note or ""):
        d = float(m.group(3)) * (1e6 if m.group(4) == "M" else 1e3)
        t = m.group(1).strip().title()
        for w in (" De ", " Los ", " La ", " A ", " Del "):
            t = t.replace(w, w.lower())
        out.append((t, d))
    return out


# --------------------------------------------------------------------------
# razón de un driver
# --------------------------------------------------------------------------

def note_fk(v):
    v = abs(v)
    return f"${v / 1e6:.1f}m" if v >= 95e3 else f"${v / 1e3:.0f}k"


def _note_key(name):
    k = _plain(note_english(name))
    return " ".join(k.split()[:3])


def _note_group(its):
    g = {}
    for name, v, seg in its:
        en = note_english(name, seg)
        if not en:
            continue
        k = " ".join(_plain(en).split()[:3])
        rec = g.setdefault(k, {"name": en, "v": 0.0, "recovery": bool(re.search(RECOVERY_RX, _plain(seg))), "units": None})
        um = re.search(r"\ben\s+(\d[\d,]*)\s+(unidades|units|vins)\b", seg, re.I)
        if um and not rec["units"]:
            rec["units"] = um.group(1)
        rec["v"] += v
    return g


def note_reason(note_m, note_p, d, is_expense, comp, prev_month_name, label=None, line_name=None):
    """Texto corto en inglés (o None) que explica la variación d de un renglón."""
    if abs(d) < 0.05e6:
        return None
    up = d < 0 if is_expense else d > 0  # True = el renglón 'subió' (más gasto / más ingreso)

    if comp == "mom":
        hb = [(h, x) for h, x in note_hubs(note_m) if x * d > 0 and abs(x) >= 0.2 * abs(d)]
        if hb:
            hb = sorted(hb, key=lambda t: -abs(t[1]))[:2]
            names = " and ".join(h for h, _ in hb)
            return f"mainly in {names} hub{'s' if len(hb) > 1 else ''}"
        var = note_variations(note_m)
        if var:
            dirn = [(n, v) for n, v in var if (v > 0) == up]
            opp = [(n, v) for n, v in var if (v > 0) != up]
            dirn = sorted(dirn, key=lambda t: -abs(t[1]))[:2]
            opp = sorted(opp, key=lambda t: -abs(t[1]))[:1]
            if dirn:
                return f"on {'higher' if up else 'lower'} " + " and ".join(f"{note_english(n)} ({note_fk(v)})" for n, v in dirn)
        cur, prv = _note_group(note_items(note_m, label)), _note_group(note_items(note_p, label))
        if cur and not note_p:  # sin nota del mes anterior no se sabe qué cambió: solo lo que se pagó
            top = [r for r in sorted(cur.values(), key=lambda r: -r["v"]) if not r["recovery"] and r["v"] <= 2.5 * abs(d)
                   and r["v"] >= 0.25 * abs(d)][:2] if up else []
            return ("mainly " + " and ".join(f"{r['name']} ({note_fk(r['v'])})" for r in top)) if top else None
        if cur and prv:
            cands = []
            for k in set(cur) | set(prv):
                a, b = cur.get(k), prv.get(k)
                rec = (a or b)["recovery"]
                change = (a["v"] if a else 0.0) - (b["v"] if b else 0.0)
                effect = -change if rec else change  # efecto en el renglón (gasto/ingreso)
                if (effect > 0) == up and abs(effect) >= 0.25 * abs(d):
                    cands.append((abs(effect), k, a, b, rec))
            cands.sort(reverse=True)
            if cands:
                _, k, a, b, rec = cands[0]
                name = (a or b)["name"]
                if a and not b:
                    return f"from {name} ({(a['units'] + ' units; ') if a.get('units') else ''}{note_fk(a['v'])})"
                if b and not a:
                    return f"as {prev_month_name} included {name} ({note_fk(b['v'])})"
                if min(a["v"], b["v"]) < 0.25 * max(a["v"], b["v"]):  # formatos distintos entre meses: no se compara
                    return f"from {name} ({note_fk(a['v'])})" if a["v"] > b["v"] else f"as {prev_month_name} included {name} ({note_fk(b['v'])})"
                word = "higher" if a["v"] > b["v"] else "lower"
                return f"on {word} {re.sub(r'^the ', '', name)} ({note_fk(a['v'])} vs. {note_fk(b['v'])})"
        if not note_m and note_p and not up:
            tot = re.search(r"(?:neto|net)\s*:?\s*\$?\s*([\d,]+(?:\.\d+)?)", note_p, re.I) or \
                re.search(r"total\s*:?\s*\$?\s*([\d,]+(?:\.\d+)?)", note_p, re.I)
            if tot:
                return f"as {prev_month_name} included {note_fk(float(tot.group(1).replace(',', '')))} of {line_name or 'these expenses'}"
        return None

    # vs. Budget: lo que la nota del mes dice que se pagó / se recuperó
    cur = _note_group(note_items(note_m, label))
    if not cur or note_variations(note_m):
        return None
    if not up and not is_expense:
        return None
    if is_expense and not up:
        rec = sorted([r for r in cur.values() if r["recovery"]], key=lambda r: -r["v"])
        if rec:
            return f"on account of {rec[0]['name']} ({note_fk(rec[0]['v'])})"
        return None
    top = sorted([r for r in cur.values() if not r["recovery"]], key=lambda r: -r["v"])
    top = [r for r in top if 0.15 * abs(d) <= r["v"] <= 2.0 * abs(d)][:2]
    if not top:
        return None
    return "mainly from " + " and ".join(f"{r['name']} ({(r['units'] + ' units; ') if r.get('units') else ''}{note_fk(r['v'])})" for r in top)


def note_month_name(period):
    return MONTHS_EN[int(period[5:7]) - 1]


# ====================================================================================================
# motor MFR: Board View -> boxes Budget/MoM
# ====================================================================================================

MINUS = "‑"  # guion no separable, como en el deck ("‑$4.6m")
TOL_ABS = 2.0

BOX_PATTERNS = [
    ("Revenues", r"^revenues?$"),
    ("Gross Profit", r"^gross profit$"),
    ("EBITDA", r"^ebitda$"),
    ("Net Income", r"^net income$"),
]

# alias de nombres como se escriben en los comentarios del MFR
DISPLAY = [
    (r"^ride hailing revenue", "ride-hailing revenues"),
    (r"^advertising revenue", "advertising revenues"),
    (r"^uber bonus", "Uber bonuses"),
    (r"^revenues?$", "Revenues"),
    (r"^opex$", "OPEX"),
    (r"^gross profit$", "GP"),
    (r"^uber take rate", "Uber TR"),
    (r"^driver payments", "driver payments"),
    (r"^ev fuel cost", "EV fuel cost"),
    (r"^charging infra", "charging infrastructure"),
    (r"^chaas", "CHaaS"),
    (r"^vehicle leasing", "vehicle leasing"),
    (r"^driving penalties", "driving penalties"),
    (r"^payroll expense", "payroll expense"),
    (r"^legal & professional", "legal & professional fees"),
    (r"^direct line sg&a|^direct business line", "DL expenses"),
    (r"^cf & corp", "CF & Corp. G&A"),
    (r"^d&a$", "D&A"),
    (r"^interest expense", "net interest expense"),
    (r"^taxes$", "taxes"),
    (r"^ebitda$", "EBITDA"),
    (r"^gain / \(loss\) from secondary sales", "loss from secondary sales"),
    (r"^other c&e", "other C&E"),
    (r"^outside services", "outside services (cleaning)"),
    (r"^others$", "other expenses"),
]

# parámetros de estilo: los ajusta vemo_mfr_learn contra tus comentarios publicados (models/mfr_style.json)
STYLE_DEFAULT = {
    "top": {"min_abs": 0.05e6, "cover": 0.97, "max_n": 4},
    "sub": {"min_abs": 0.1e6, "cover": 0.85, "max_n": 7, "expand_min": 0.3e6},
    "offs": {"min_abs": 0.1e6, "rel": 0.03, "max_n": 3},
    "resid": {"on": True, "min_abs": 0.15e6, "rel": 0.05},
    "rev": {"min_abs": 0.1e6, "cover": 0.9, "max_n": 3},
    "display": {},          # cómo nombras cada renglón (aprendido de tus decks)
    "expand_word": {"budget": "explained by:", "mom": "explained by:"},
    "always_expand": [],    # renglones que siempre desglosas (p. ej. OPEX, DL expenses)
    "no_expand": [],        # renglones que nunca desglosas en (a)(b) (p. ej. CF & Corp. G&A)
    "ni_mom_inline": False, # Net Income MoM = EBITDA MoM -> "In line."
    "kpi_learned": {},      # renglón -> KPI que citas "(x vs. y)" y tu frase
    "recurring": {},        # renglón|dirección -> razón que repites mes a mes (se marca para confirmar)
    "slightly_pct": 0.0,    # "slightly higher SH" cuando el cambio es menor a este %
    "per_box": {},          # ajustes propios de cada box (GP, EBITDA, Net Income) cuando escribes distinto en cada uno
}
STYLE = json.loads(json.dumps(STYLE_DEFAULT))


# Perfil "calidad": cuánto detalle y qué tan completo es cada box. Es el estándar de los boxes de referencia
# (eval/oro_calidad.json): todos los drivers que mueven la aguja, con su porqué, sumando al encabezado.
# El vocabulario (cómo nombras cada renglón, tus conectores, tus KPIs) sigue saliendo de tus decks.
QUALITY = {
    "top": {"min_abs": 0.05e6, "cover": 1.0, "max_n": 6},
    "sub": {"min_abs": 0.15e6, "rel_gross": 0.035, "cover": 1.0, "max_n": 7, "expand_min": 0.5e6},
    "offs": {"min_abs": 0.15e6, "rel": 0.0, "rel_gross": 0.035, "max_n": 4},
    "resid": {"on": True, "min_abs": 0.15e6, "rel": 0.0},
    "rev": {"min_abs": 0.3e6, "cover": 0.97, "max_n": 3},
    "always_expand": ["opex", "direct line sg&a", "cf & corp. g&a"],
    "no_expand": ["d&a", "taxes", "interest expense", "revenue", "gross profit"],
    "ni_mom_inline": True,
    "recurring_mode": "sugerir",
    "kpi_min_pct": 2.0,
    "top_offs": {"min_abs": 0.05e6, "rel": 0.0, "rel_gross": 0.0, "max_n": 6},
    "per_box": {},
}


def load_style(path=None, perfil="calidad"):
    """Carga el estilo aprendido de tus decks; con perfil 'calidad' (default) el nivel de detalle es el de
    referencia y de tus decks se toma el vocabulario. perfil='decks' imita también cuánto detallas tú."""
    global STYLE
    st = json.loads(json.dumps(STYLE_DEFAULT))
    try:
        if path is None:
            path = os.path.join(MODELS_DIR, "mfr_style.json")
        if path and os.path.exists(path):
            with open(path, encoding="utf-8") as fh:
                learned = json.load(fh).get("style", {})
            for k, v in learned.items():
                if isinstance(v, dict) and isinstance(st.get(k), dict):
                    st[k].update(v)
                else:
                    st[k] = v
    except Exception:
        pass
    if perfil == "calidad":
        for k, v in QUALITY.items():
            st[k] = json.loads(json.dumps(v))
    STYLE = st
    return st


def set_style(st):
    global STYLE
    STYLE = st


MERGES = [("CF & Corp. G&A", r"^cross-?\s?functional|^corporate (expenses|bonus|g&a)")]

# KPIs que acompañan a una línea (renglón de KPI buscado por etiqueta en la misma hoja)
KPI_HOOKS = {
    r"^uber take rate": [("Financial Uber Take Rate %", "pct", None)],
    r"^payroll expense": [("Headcount", "int", "HC")],
    r"^maintenance$": [("Maintenance Per km", "money2", "per km")],
    r"^washing$": [("Daily Washing Expense per Vehicle", "money1", "per vehicle per day")],
    r"^telematics$": [("Monthly Telematics per Vehicle", "money0", "per vehicle")],
    r"^charging infra": [("Expense per Hub", "money0", "per hub")],
    r"^security$": [("Security Expense per Hub", "money0", "per hub")],
}
# razones calculadas con tus KPIs: costo por unidad (renglón / KPI) y tamaño de la flota
DERIVED_HOOKS = {
    r"^ev fuel cost": ("operational vehicles", "vehicles"),
}

RESID_NAME = {"opex": "other fleet operating expenses", "direct line sg&a": "other G&A expenses",
              "cf & corp. g&a": "other corporate expenses", "revenue": "other revenues"}


# --------------------------------------------------------------------------
# lectura de la hoja
# --------------------------------------------------------------------------

def _plain(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def _num(x):
    return float(x) if isinstance(x, (int, float)) and not isinstance(x, bool) and math.isfinite(x) else None


def _ym(d):
    return f"{d.year:04d}-{d.month:02d}" if isinstance(d, (_dt.date, _dt.datetime)) else None


def find_board_sheet(wb, company=None):
    names = [n for n in wb.sheetnames if "board view (actuals)" in n.lower()]
    if company:
        pref = [n for n in names if n.lower().startswith(company.lower())]
        names = pref or names
    if not names:
        raise RuntimeError("no encontré la hoja 'Board View (Actuals)'")
    return wb[names[0]]


def _layout(ws, period):
    """Columnas: serie mensual (fecha->col), bloque MoM (M-1, M) y bloque Budget (M, Budget)."""
    maxc = min(ws.max_column, 200)
    date_row = None
    for r in range(1, 20):
        n = sum(isinstance(ws.cell(r, c).value, (_dt.date, _dt.datetime)) for c in range(1, maxc + 1))
        if n >= 12:
            date_row = r
            break
    if date_row is None:
        raise RuntimeError("no encontré el renglón de fechas")
    series = {}
    for c in range(1, maxc + 1):
        ym = _ym(ws.cell(date_row, c).value)
        if ym and ym not in series:
            series[ym] = c  # la primera aparición es la serie mensual
    # bloques por título
    title_row = None
    for r in range(1, date_row):
        if any(isinstance(ws.cell(r, c).value, str) and "mom view" in ws.cell(r, c).value.lower() for c in range(1, maxc + 1)):
            title_row = r
            break
    kind_row = None
    for r in range(date_row, date_row + 10):
        vals = [str(ws.cell(r, c).value or "").strip().lower() for c in range(1, maxc + 1)]
        if vals.count("actual") >= 6 and "budget" in vals:
            kind_row = r
            break
    lay = {"date_row": date_row, "series": series, "m": None, "p": None, "a": None, "b": None, "block": False}
    if title_row and kind_row:
        titles = {c: str(ws.cell(title_row, c).value or "").lower() for c in range(1, maxc + 1)}
        mom_c = next((c for c, t in titles.items() if "mom view" in t), None)
        bud_c = next((c for c, t in titles.items() if "month actual vs. budget" in t or "month actual vs budget" in t), None)
        if mom_c and _ym(ws.cell(date_row, mom_c + 1).value) == period:
            lay["p"], lay["m"] = mom_c, mom_c + 1
        if bud_c:
            for c in range(bud_c, bud_c + 4):
                if str(ws.cell(kind_row, c).value or "").strip().lower() == "budget" and _ym(ws.cell(date_row, c).value) == period:
                    lay["b"] = c
                    lay["a"] = c - 1
                    break
    prev = _prev(period)
    # Actual y M-1 siempre de la serie mensual: es la única donde cada total = suma de sus renglones
    # (el bloque "MoM View" omite algunos renglones, p. ej. Severance). El Budget sale del bloque.
    lay["block"] = lay["b"] is not None
    lay["m"], lay["p"] = series.get(period), series.get(prev)
    if lay["m"] is None:
        raise RuntimeError(f"la hoja no tiene la columna de {period}")
    lay["note_m"], lay["note_p"] = series.get(period), series.get(prev)
    return lay


def _prev(period):
    y, m = int(period[:4]), int(period[5:7])
    return f"{y - 1:04d}-12" if m == 1 else f"{y:04d}-{m - 1:02d}"


def _label(ws, r):
    for c in range(2, 17):
        v = ws.cell(r, c).value
        if isinstance(v, str) and v.strip():
            return v.strip(), c
    return None, None


def _note(ws, r, c):
    if not c:
        return None
    cm = ws.cell(r, c).comment
    if not cm or not cm.text:
        return None
    t = re.sub(r"^\s*(vemo|VEMO)\s*:\s*", "", cm.text.strip())
    t = re.sub(r"[ \t]*\n[ \t]*", " | ", t)
    t = re.sub(r"(\s*\|\s*)+", " | ", t).strip(" |")
    return t or None


def _budget_lookup(wb, ws, period):
    """Mes histórico: el Budget sale de la hoja '<EMPRESA> Board View (Budget)' (misma etiqueta y orden)."""
    if wb is None:
        return None
    co = ws.title.split()[0].lower()
    cands = [n for n in wb.sheetnames if "board view (budget" in n.lower() and n.lower().startswith(co)]
    for name in cands:
        bs = wb[name]
        maxc = min(bs.max_column, 200)
        date_row = kind_row = None
        for r in range(1, 20):
            if sum(isinstance(bs.cell(r, c).value, (_dt.date, _dt.datetime)) for c in range(1, maxc + 1)) >= 12:
                date_row = r
                break
        if not date_row:
            continue
        for r in range(date_row, date_row + 10):
            if sum(str(bs.cell(r, c).value or "").strip().lower() == "budget" for c in range(1, maxc + 1)) >= 6:
                kind_row = r
                break
        col = None
        for c in range(1, maxc + 1):
            if _ym(bs.cell(date_row, c).value) == period and \
                    (kind_row is None or str(bs.cell(kind_row, c).value or "").strip().lower() == "budget"):
                col = c  # sin renglón Actual/Budget: toda la hoja es budget
                break
        if not col:
            continue
        occ, table = {}, {}
        for r in range(1, bs.max_row + 1):
            lab, _ = _label(bs, r)
            if not lab:
                continue
            k = _plain(lab)
            occ[k] = occ.get(k, 0) + 1
            table[(k, occ[k])] = _num(bs.cell(r, col).value)
        return table
    return None


def _col_letter(c):
    out = ""
    while c:
        c, r = divmod(c - 1, 26)
        out = chr(65 + r) + out
    return out


def read_board(ws, period, wb=None):
    lay = _layout(ws, period)
    btab = _budget_lookup(wb, ws, period) if lay["b"] is None else None
    occ = {}
    rows, kpis = [], {}
    start = end = None
    for r in range(1, ws.max_row + 1):
        lab, _ = _label(ws, r)
        if lab and start is None and re.search(r"p&l", lab.lower()):
            start = r
        if start and lab and r > start and _plain(lab) == "net income":
            end = r
            break
    if not start or not end:
        raise RuntimeError("no encontré la sección P&L (… hasta Net Income)")
    for r in range(1, ws.max_row + 1):
        lab, col = _label(ws, r)
        if not lab:
            continue
        occ[_plain(lab)] = occ.get(_plain(lab), 0) + 1
        rec = {"row": r, "label": lab, "key": _plain(lab),
               "m": _num(ws.cell(r, lay["m"]).value), "p": _num(ws.cell(r, lay["p"]).value) if lay["p"] else None,
               "b": _num(ws.cell(r, lay["b"]).value) if lay["b"] else None,
               "note_m": _note(ws, r, lay["note_m"]), "note_p": _note(ws, r, lay["note_p"]),
               "hist": [_num(ws.cell(r, c).value) for ym, c in sorted(lay["series"].items()) if ym <= period][-4:],
               "series": {ym: _num(ws.cell(r, c).value) for ym, c in lay["series"].items() if ym <= period},
               "occ": occ[_plain(lab)], "cell": f"{_col_letter(lay['m'])}{r}"}
        if btab is not None:
            rec["b"] = btab.get((rec["key"], occ[rec["key"]]))
        if start < r <= end:
            if re.search(r"margin|%", rec["key"]):
                continue
            hist = {}
            for ym, c in lay["series"].items():
                if ym <= period:
                    t = _note(ws, r, c)
                    if t:
                        hist[ym] = t
            rec["notes_hist"] = hist
            rows.append(rec)
        elif r < start and rec["key"] not in kpis:
            kpis[rec["key"]] = rec
    days = kpis.get("days per month")
    return {"layout": lay, "rows": rows, "kpis": kpis, "days": days}


# --------------------------------------------------------------------------
# árbol del P&L
# --------------------------------------------------------------------------

def _vec(n):
    return [n["m"] or 0.0, n["p"] or 0.0] + [h or 0.0 for h in n["hist"]]


def _eq(a, b):
    return all(abs(x - y) <= max(TOL_ABS, 1e-6 * max(abs(x), abs(y))) for x, y in zip(a, b))


def build_tree(rows):
    stack = []
    for n in rows:
        n["children"] = []
        v = _vec(n)
        if any(abs(x) > TOL_ABS for x in v):
            for k in range(len(stack), 1, -1):
                block = stack[-k:]
                s = [sum(col) for col in zip(*[_vec(x) for x in block])]
                if _eq(s, v):
                    n["children"] = block
                    del stack[-k:]
                    break
        stack.append(n)
    return stack


def _find(nodes, rx):
    out = []

    def walk(ns):
        for x in ns:
            if re.search(rx, x["key"]):
                out.append(x)
            walk(x["children"])
    walk(nodes)
    return out[0] if out else None


# --------------------------------------------------------------------------
# variaciones
# --------------------------------------------------------------------------

def _ref(n, comp):
    return n["b"] if comp == "budget" else n["p"]


def delta(n, comp):
    if n["m"] is None and _ref(n, comp) is None:
        return 0.0
    return (n["m"] or 0.0) - (_ref(n, comp) or 0.0)


def budget_gap(n):
    """Diferencia entre el budget del renglón y la suma del budget de sus hijos (None si no aplica)."""
    if not n["children"] or n["b"] is None:
        return None
    return n["b"] - sum((c["b"] or 0.0) for c in n["children"])


def _decomposable(n, comp):
    if not n["children"]:
        return False
    if comp == "mom":
        return True
    gap = budget_gap(n)
    if gap is None:
        return False
    # tolerancia: un descuadre chico del archivo no impide explicar, pero se reporta en "Revisar"
    return abs(gap) <= max(TOL_ABS, 1e-6 * abs(n["b"])) or abs(gap) <= max(0.15e6, 0.2 * abs(delta(n, comp)))


EXPENSE_RX = r"expense|^taxes$|^d&a$|^opex$|sg&a|g&a|fees|payments|cost|take rate|leasing|insurance|maintenance|penalt"


def _is_expense(n):
    if re.search(EXPENSE_RX, n["key"]) and not re.search(r"profit|revenue", n["key"]):
        return True
    vals = [x for x in [n["m"], n["p"], n["b"]] + n["hist"] if x]
    return bool(vals) and sorted(vals)[len(vals) // 2] < 0 and not re.search(r"profit|ebitda|ebit|ebt|net income|revenue", n["key"])


def _display(label):
    k = _plain(label)
    learned = STYLE.get("display", {}).get(k)
    if learned:
        keep = {"Uber", "TR", "GP", "DL", "CF", "G&A", "D&A", "OPEX", "EV", "CHaaS", "LTO", "HC", "SG&A", "Corp.", "RPH", "SH", "PPR"}
        return " ".join(w if (w in keep or w.isupper()) else w.lower() for w in learned.split())
    for rx, name in DISPLAY:
        if re.search(rx, k):
            return name
    words = label.strip().split()
    return " ".join(w if (w.isupper() or "&" in w or any(ch.isdigit() for ch in w)) else w.lower() for w in words)


def _word(n, d):
    if _is_expense(n):
        return "higher" if d < 0 else "lower"
    return "higher" if d > 0 else "lower"


def fm(x, ca=False):
    """+$1.2m / ‑$0.4m (en millones, 1 decimal)."""
    v = x / 1e6
    s = f"{abs(v):.1f}"
    sign = "+" if v >= 0 else MINUS
    if s == "0.0":
        sign = "+" if v >= 0 else MINUS
    return f"{'ca. ' if ca else ''}{sign}${s}m"


def _virtual(label, parts):
    def sm(f):
        vals = [p[f] for p in parts]
        return None if all(v is None for v in vals) else sum(v or 0.0 for v in vals)
    return {"row": parts[0]["row"], "label": label, "key": _plain(label), "m": sm("m"), "p": sm("p"), "b": sm("b"),
            "note_m": " / ".join(p["note_m"] for p in parts if p["note_m"]) or None,
            "note_p": " / ".join(p["note_p"] for p in parts if p["note_p"]) or None,
            "hist": [sum((p["hist"][i] if i < len(p["hist"]) else 0.0) or 0.0 for p in parts) for i in range(4)],
            "children": parts}


def deck_children(n, box_keys):
    """Hijos tal como se leen en la tabla del deck: se aplanan subtotales intermedios (EBIT, EBT, SG&A,
    Other G&A) y se juntan Cross-functional + Corporate en 'CF & Corp. G&A'."""
    out = []
    for c in n["children"]:
        if c["children"] and c["key"] not in box_keys and re.search(r"^(ebit|ebt|sg&a|other g&a)$", c["key"]):
            out.extend(deck_children(c, box_keys))
        else:
            out.append(c)
    for name, rx in MERGES:
        grp = [c for c in out if re.search(rx, c["key"])]
        if len(grp) >= 2:
            i = out.index(grp[0])
            out = [c for c in out if c not in grp]
            out.insert(i, _virtual(name, grp))
    return out


def _flat_sub(n):
    """Sub-drivers (a)(b)(c): hijos del renglón, aplanando 'Other G&A'."""
    out = []
    for c in n["children"]:
        if c["children"] and re.search(r"^other g&a$", c["key"]):
            out.extend(_flat_sub(c))
        else:
            out.append(c)
    return out


def _select(items, total, comp, min_abs=0.05e6, cover=0.85, max_n=5, rel_gross=0.0, offs_cfg=None, **_):
    same = sorted([x for x in items if delta(x, comp) * total > 0], key=lambda x: -abs(delta(x, comp)))
    opp = sorted([x for x in items if delta(x, comp) * total < 0], key=lambda x: -abs(delta(x, comp)))
    gross = sum(abs(delta(x, comp)) for x in same) or 1.0
    moved = sum(abs(delta(x, comp)) for x in items)  # movimiento bruto: lo que se movió de verdad, en ambos sentidos
    min_abs = max(min_abs, rel_gross * moved)
    chosen, acc = [], 0.0
    for x in same:
        d = abs(delta(x, comp))
        if d < min_abs or len(chosen) >= max_n:
            break
        chosen.append(x)
        acc += d
        if acc / gross >= cover and len(chosen) >= 1:
            break
    o = offs_cfg or STYLE["offs"]
    thr = max(o["min_abs"], o["rel"] * abs(total), o.get("rel_gross", 0.0) * moved)
    offs = [x for x in opp if abs(delta(x, comp)) >= thr][:int(o["max_n"])]
    return chosen, offs


# --------------------------------------------------------------------------
# KPIs
# --------------------------------------------------------------------------

def _kv(k, comp):
    if not k:
        return None, None
    return k["m"], (k["b"] if comp == "budget" else k["p"])


def _fmt_kpi(v, kind):
    if v is None:
        return "n.d."
    if kind == "pct":
        return f"{v * 100:.1f}%"
    if kind == "int":
        return f"{v:,.0f}"
    if kind == "money1":
        return f"${abs(v):,.1f}"
    if kind == "money2":
        return f"${abs(v):,.2f}"
    if kind == "money0":
        return f"${abs(v) / 1000:,.0f}k" if abs(v) >= 10_000 else f"${abs(v):,.0f}"
    if kind == "k":
        return f"{v / 1000:,.1f}k"
    return f"{v:,.1f}"


def _pct(a, b):
    if a is None or not b:
        return None
    return (a / b - 1) * 100


def _pstr(p):
    return f"{'+' if p >= 0 else MINUS}{abs(p):.1f}%"


NOTE_FACTS = {  # cifras operativas que FP&A deja en la nota de celda (se leen, no se inventan)
    r"^vehicle leasing": [r"unidades con leasing:\s*(\d[\d,]*)", r"aumento a (\d[\d,]*)", r"(\d[\d,]*)\s*(?:u\b|unidades)"],
}


def _note_num(text, rxs):
    if not text:
        return None
    for rx in rxs:
        m = re.search(rx, text, re.I)
        if m:
            return int(m.group(1).replace(",", ""))
    return None


def note_suffix(n, comp):
    for rx, rxs in NOTE_FACTS.items():
        if re.search(rx, n["key"]):
            a = _note_num(n.get("note_m"), rxs)
            if a is None:
                return ""
            if comp == "mom":
                b = _note_num(n.get("note_p"), rxs)
                return f"{a:,} vs. {b:,} vehicles" if b else f"{a:,} vehicles"
            return f"{a:,} vehicles" + ("; not budgeted" if n["b"] is None else "")
    return "not budgeted" if comp == "budget" and n["b"] is None and n["m"] else ""


def kpi_suffix(n, kpis, comp):
    ns = note_suffix(n, comp)
    ks = _kpi_suffix(n, kpis, comp)
    return "; ".join(x for x in (ks, ns) if x)


def _derived_suffix(n, kpis, comp):
    for rx, (klab, unit) in DERIVED_HOOKS.items():
        if not re.search(rx, n["key"]):
            continue
        a, r = _kv(kpis.get(klab), comp)
        ref = _ref(n, comp)
        if not a or not r or not n["m"] or not ref:
            return ""
        ua, ur = abs(n["m"]) / a, abs(ref) / r
        parts = []
        if abs(ua / ur - 1) >= 0.03:
            parts.append(f"${ua / 1000:,.1f}k vs. ${ur / 1000:,.1f}k per vehicle")
        if abs(a / r - 1) >= 0.02:
            parts.append(f"{a:,.0f} vs. {r:,.0f} {unit}")
        return "; ".join(parts)
    return ""


def _kpi_suffix(n, kpis, comp):
    dv = _derived_suffix(n, kpis, comp)
    if dv:
        return dv
    for rx, hooks in KPI_HOOKS.items():
        if re.search(rx, n["key"]):
            parts = []
            for lab, kind, unit in hooks:
                a, r = _kv(kpis.get(_plain(lab)), comp)
                if a is None or r is None or not r or abs(a / r - 1) * 100 < STYLE.get("kpi_min_pct", 0.0):
                    continue
                txt = f"{_fmt_kpi(a, kind)} vs. {_fmt_kpi(r, kind)}"
                parts.append(f"{unit} {txt}" if unit in ("HC",) else (f"{txt} {unit}" if unit else txt))
            return "; ".join(parts)
    return ""


def rph_phrase(kpis, comp, days=None):
    """Driver de ride-hailing en el lenguaje del MFR: RPH = PPR x TPH ; Revenue = RPH x SH."""
    rph, ppr, sh, tph = (kpis.get(k) for k in ("revenue per hour", "price per ride", "supply hours", "trips per hour"))
    if not (rph and ppr and sh):
        return None
    ra, rr = _kv(rph, comp)
    pa, pr = _kv(ppr, comp)
    sa, sr = _kv(sh, comp)
    if None in (ra, rr, pa, pr, sa, sr) or min(ra, rr, sa, sr) <= 0 or not pr:
        return None
    c_rph, c_sh = math.log(ra / rr), math.log(sa / sr)
    rph_txt = f"{'higher' if c_rph >= 0 else 'lower'} RPH {_pstr(_pct(ra, rr))} (${ra:,.1f} vs. ${rr:,.1f})"
    ppr_txt = f"{'higher' if pa >= pr else 'lower'} PPR {_pstr(_pct(pa, pr))} (${pa:,.1f} vs. ${pr:,.1f})"
    sl = "slightly " if abs(_pct(sa, sr)) < STYLE.get("slightly_pct", 0) else ""
    sh_txt = f"{sl}{'higher' if sa >= sr else 'lower'} SH {_pstr(_pct(sa, sr))} ({sa / 1000:,.1f}k vs. {sr / 1000:,.1f}k)"
    if comp == "mom" and days and days.get("m") and days.get("p") and days["m"] != days["p"]:
        n = abs(days["m"] - days["p"])
        sh_txt += f", on account of {'one' if n == 1 else f'{n:.0f}'} {'more' if days['m'] > days['p'] else 'less' if n == 1 else 'fewer'} calendar day{'s' if n != 1 else ''} ({days['m']:.0f} vs. {days['p']:.0f})"
    rph_full = f"{rph_txt}, on account of a {ppr_txt}"
    return {"rph": rph_full, "sh": sh_txt, "rph_first": abs(c_rph) >= abs(c_sh), "same": (c_rph >= 0) == (c_sh >= 0),
            "c_rph": c_rph, "c_sh": c_sh}


# --------------------------------------------------------------------------
# redacción de un box
# --------------------------------------------------------------------------

ROMAN = ["i", "ii", "iii", "iv", "v", "vi", "vii", "viii"]
MONTHS_EN = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
             "November", "December"]
ALPHA = "abcdefghjklmn"  # sin "i" para no confundir con el romano (i)


def _alloc(vals, target):
    """Redondeo a décimas que SUMA exacto `target` (décimas): mayor residuo. vals en pesos."""
    x = [v / 1e5 for v in vals]
    r = [round(v) for v in x]
    diff = int(target - sum(r))
    step = 1 if diff > 0 else -1
    for _ in range(abs(diff)):
        i = max(range(len(x)), key=lambda j: (x[j] - r[j]) * step)
        r[i] += step
    return r


def _cap(t):
    return t[:1].upper() + t[1:]


def _item(n, comp, kpis, with_kpi=True, deep=False, ctx=None, sub=False):
    d = delta(n, comp)
    k = kpi_suffix(n, kpis, comp) if with_kpi else ""
    dv = (ctx or {}).get("disp", {}).get((id(n), comp))
    amt = fm(dv * 1e5) if dv is not None else fm(d)
    unb = comp == "budget" and n["b"] is None and bool(n["m"])
    k = k.replace("; not budgeted", "").replace("not budgeted", "").strip("; ")
    s = f"{_word(n, d)} {_display(n['label'])} ({amt}{'; ' + k if k else ''})"
    why = None
    learned = STYLE.get("kpi_learned", {}).get(n["key"])
    if learned:
        k = re.sub(r"(;\s*)?(HC [\d,]+ vs\. [\d,]+|[\d,]+ vs\. [\d,]+ active drivers)", "", k).strip("; ")
        s = s.replace(f"({amt}; {kpi_suffix(n, kpis, comp)})", f"({amt}{'; ' + k if k else ''})") if not unb else s
        why = _learned_kpi(learned, kpis, comp, line_up=(d < 0) == _is_expense(n))
    if why:
        pass
    elif re.search(r"^driver payments", n["key"]):
        a, r = _kv(kpis.get("eop active drivers"), comp)
        if a and r and abs(a - r) >= 1:
            more = a > r
            if more == (d < 0):
                why = f"{'due to additional' if more else 'on fewer'} drivers ({a:,.0f} vs. {r:,.0f})"
    elif not any(re.search(rx, n["key"]) for rx in NOTE_FACTS):
        prev_name = note_month_name(ctx["prev_period"]) if ctx else "the prior month"
        why = note_reason(n.get("note_m"), n.get("note_p"), d, _is_expense(n), comp, prev_name, n["label"], _display(n["label"]))
    if not why:
        dirn = "up" if (d < 0) == _is_expense(n) else "down"
        rec = STYLE.get("recurring", {}).get(f"{n['key']}|{comp}|{dirn}") or STYLE.get("recurring", {}).get(f"{n['key']}|*|{dirn}")
        if rec:
            if STYLE.get("recurring_mode") == "sugerir":
                if ctx is not None:
                    ctx["flags"].append(f"{_cap(_display(n['label']))} ({amt}): sin nota este mes; en meses anteriores lo explicaste "
                                        f"\"{rec}\". Si aplica, agrégalo.")
            else:
                why = rec
                if ctx is not None:
                    ctx["flags"].append(f"{_cap(_display(n['label']))}: usé tu razón de meses anteriores (\"{rec}\"); confírmala para este mes.")
    if why:
        s += f", {why}" if not why.startswith(("from ", "due ", "on ")) or "," in s[-3:] else f" {why}"
    if deep and n["children"] and _decomposable(n, comp) and abs(d) >= 0.2e6:
        top = max(n["children"], key=lambda c: abs(delta(c, comp)))
        td = delta(top, comp)
        top_ref = _ref(top, comp)
        released = _is_expense(n) and (top["m"] or 0) > 0 and (top_ref is None or top_ref <= 0)
        if td * d > 0 and released and abs(td) >= 0.4 * abs(d):
            mon = MONTHS_EN[int(ctx["period"][5:7]) - 1] if ctx and ctx.get("period") else "the month"
            s += f", due to a release of provisions in the {_display(top['label'])} line ({fm(top['m'])} in {mon})"
            if ctx is not None:
                ctx["flags"].append(f"{_cap(_display(top['label']))} con saldo a favor ({fm(top['m'])}): lo redacté como liberación de "
                                    f"provisiones; confirma (y si es de 2025, agrega 'unused 2025').")
            why = why or "release"
        elif td * d > 0 and abs(td) >= 0.6 * abs(d) and abs(td) >= 0.1e6:
            s += f", mainly from the {_display(top['label'])} line ({fm(td)})"
            if ctx is not None and not top.get("note_m"):
                ctx["flags"].append(f"{_cap(_display(top['label']))} dentro de {_display(n['label'])} ({fm(td)} {'vs. Budget' if comp == 'budget' else 'MoM'}): sin nota en el Excel; confirma la causa.")
            why = why or "line"
    if ctx is not None:
        if unb and abs(n["m"] or 0) >= 0.5e6:
            msg = f"{_cap(_display(n['label']))} ({fm(n['m'])}) no tiene renglón de budget: confirma si está presupuestado dentro de otro renglón."
            if msg not in ctx["flags"]:
                ctx["flags"].append(msg)
        elif sub and abs(d) >= 0.5e6 and not why and not k and not n["children"]:
            ctx["flags"].append(f"{_cap(_display(n['label']))} ({amt} {'vs. Budget' if comp == 'budget' else 'MoM'}): sin nota del mes que explique la variación.")
    return s


def _fmt_learned(v, kind):
    if kind == "pct":
        return f"{v * 100:.1f}%" if abs(v) < 1.5 else f"{v:.1f}%"
    if kind == "money":
        return f"${v:,.2f}" if abs(v) < 10 else f"${v:,.1f}"
    if kind == "k":
        return f"{v / 1000:,.1f}k"
    return f"{v:,.0f}"


def _learned_kpi(hooks, kpis, comp, line_up=None):
    for h in hooks:
        a, r = _kv(kpis.get(h["kpi"]), comp)
        if a is None or r is None or a == r:
            continue
        if r and abs(a / r - 1) * 100 < STYLE.get("kpi_min_pct", 0.0):  # 158 vs. 160 no explica nada
            continue
        if line_up is not None and (a > r) != line_up:  # "lower payroll due to a higher HC" no tiene sentido
            continue
        phrase = h["up"] if a > r else h["down"]
        return f"{phrase} ({_fmt_learned(a, h.get('fmt'))} vs. {_fmt_learned(r, h.get('fmt'))})"
    return None


def _rd(x):
    """Monto como se lee en el box (millones, 1 decimal): para que los drivers sumen el encabezado."""
    return round(round(x / 1e5) / 10, 1)


def _join(parts, sep="; "):
    if len(parts) <= 1:
        return "".join(parts)
    if len(parts) == 2:
        return parts[0] + ", and " + parts[1]
    return sep.join(parts[:-1]) + sep + "and " + parts[-1]


def compose_box(box, node, comp, ctx):
    """Aplica los ajustes aprendidos para ESTE box (si hay) y redacta."""
    global STYLE
    over = (STYLE.get("per_box") or {}).get(box)
    if not over:
        return _compose_box(box, node, comp, ctx)
    base = STYLE
    merged = json.loads(json.dumps(base))
    for k, v in over.items():
        if isinstance(v, dict) and isinstance(merged.get(k), dict):
            merged[k].update(v)
        else:
            merged[k] = v
    STYLE = merged
    try:
        return _compose_box(box, node, comp, ctx)
    finally:
        STYLE = base


def _compose_box(box, node, comp, ctx):
    total = delta(node, comp)
    head = f"{'Budget' if comp == 'budget' else 'MoM'} (MXN {fm(total)})"
    evid, used = [], []
    if abs(total) < 0.05e6:
        return {"text": f"{head}: In line.", "evidence": evid, "check": None}
    if comp == "budget" and node["b"] is None:
        return {"text": f"{head}: (sin presupuesto en el archivo)", "evidence": evid, "check": None}
    kids = deck_children(node, ctx["box_keys"]) if _decomposable(node, comp) else []
    if not kids:
        return {"text": f"{head}: (sin desglose en el archivo)", "evidence": evid, "check": None}

    if box == "Revenues":
        return _compose_revenue(head, node, kids, total, comp, ctx)
    if box == "Net Income" and comp == "mom" and STYLE.get("ni_mom_inline"):
        eb = next((k for k in kids if k["key"] == "ebitda"), None)
        if eb is not None and abs(total - delta(eb, comp)) < 0.15e6:
            return {"text": f"{head}: In line.", "evidence": evid, "check": None, "items": [(eb["label"], delta(eb, comp), 1)]}

    main, offs = _select(kids, total, comp, offs_cfg=STYLE.get("top_offs"), **STYLE["top"])
    items_out = []
    disp = ctx.setdefault("disp", {})
    listed = main + offs
    tgt = round(total / 1e5) if abs(total - sum(delta(x, comp) for x in listed)) < 0.05e6 else \
        round(sum(delta(x, comp) for x in listed) / 1e5)
    for x, t in zip(listed, _alloc([delta(x, comp) for x in listed], tgt)):
        disp[(id(x), comp)] = t
    seq = main + offs
    txt, idx = [], 0
    for grp_i, grp in enumerate((main, offs)):
        chunks = []
        for n in grp:
            single = len(main) + len(offs) == 1
            s = ("" if single else f"({ROMAN[idx]}) ") + _item(n, comp, ctx["kpis"], ctx=ctx)
            idx += 1
            used.append(n)
            items_out.append((n["label"], delta(n, comp), 1))
            expandable = n["key"] not in ctx["box_keys"] and n["key"] not in STYLE.get("no_expand", []) \
                and _decomposable(n, comp) and (n["key"] in STYLE.get("always_expand", []) or
                                                max([abs(delta(n, comp))] + [abs(delta(c, comp)) for c in _flat_sub(n)])
                                                >= STYLE["sub"]["expand_min"])
            if expandable:
                subs = _flat_sub(n)
                sm, so = _select(subs, delta(n, comp), comp, **STYLE["sub"])
                items_out += [(x["label"], delta(x, comp), 2) for x in sm + so]
                # cierre: los drivers listados + "otros" suman EXACTO (en décimas) el monto que se explica
                parent_t = disp.get((id(n), comp), round(delta(n, comp) / 1e5))
                exact_rest = delta(n, comp) - sum(delta(x, comp) for x in sm + so)
                # cada driver con su redondeo natural; la diferencia (si hay) la absorbe "otros"
                nat = [round(delta(x, comp) / 1e5) for x in sm + so]
                for x, t in zip(sm + so, nat):
                    disp[(id(x), comp)] = t
                other_t = parent_t - sum(nat)
                if sm:
                    parts = [f"({ALPHA[j]}) {_item(x, comp, ctx['kpis'], deep=True, ctx=ctx, sub=True)}" for j, x in enumerate(sm)]
                    offs_t = [_item(x, comp, ctx["kpis"], ctx=ctx, sub=True) for x in so]
                    if other_t != 0:
                        rest = other_t * 1e5
                        lab = RESID_NAME.get(n["key"], f"other {_display(n['label'])}")
                        items_out.append(("other " + n["label"], rest, 2))
                        # si "otros" pesa, se dice de qué está hecho (los renglones que más lo mueven)
                        comp_txt = ""
                        if abs(rest) >= 0.3e6:
                            pool = [x for x in subs if x not in sm + so and delta(x, comp) * rest > 0]
                            pool = sorted(pool, key=lambda x: -abs(delta(x, comp)))[:3]
                            pool = [x for x in pool if abs(delta(x, comp)) >= 0.1e6]
                            if pool:
                                comp_txt = ": " + ", ".join(f"{_display(x['label'])} {fm(delta(x, comp))}" for x in pool)
                        rtxt = f"{'higher' if (rest < 0) == _is_expense(n) else 'lower'} {lab} ({fm(rest)}{comp_txt})"
                        if rest * delta(n, comp) > 0:
                            parts.append(f"({ALPHA[len(parts)]}) {rtxt}")
                        else:
                            offs_t.append(rtxt)
                    if len(parts) == 1:
                        parts = [re.sub(r"^\(a\) ", "", parts[0])]
                    s += f", {STYLE.get('expand_word', {}).get(comp, 'explained by:')} " + _join(parts)
                    if offs_t:
                        sep = "; " if any("," in t for t in offs_t) else ", "
                        s += "; partially offset by " + (sep.join(offs_t[:-1]) + sep + "and " + offs_t[-1] if len(offs_t) > 1 else offs_t[0])
                    for x in sm + so:
                        evid.append(_ev(x, comp))
                elif n.get("note_m"):
                    evid.append(_ev(n, comp))
            elif n.get("note_m") and n["key"] not in ctx["box_keys"]:
                evid.append(_ev(n, comp))
            chunks.append(s)
        if not chunks:
            continue
        if grp_i == 0:
            deep = any("(a) " in c for c in chunks)
            body = _join(chunks) if (len(chunks) > 2 or deep) else ", and ".join(chunks)
            rest_top = abs(total - sum(delta(x, comp) for x in main + offs)) >= 0.01e6
            txt.append(("Mainly driven by " if len(main) > 1 or offs or rest_top else "Driven by ") + body + ".")
        else:
            txt.append("This was partially offset by " + (_join(chunks) if len(chunks) > 1 else chunks[0]) + ".")
    covered = sum(delta(n, comp) for n in seq)
    return {"text": f"{head}: " + " ".join(txt), "evidence": [e for e in evid if e], "check": (total, covered), "items": items_out}


def _compose_revenue(head, node, kids, total, comp, ctx):
    evid = []
    rh = next((k for k in kids if re.search(r"ride hailing", k["key"])), None)
    rp = rph_phrase(ctx["kpis"], comp, ctx.get("days")) if rh else None
    others = [k for k in kids if k is not rh]
    om, oo = _select(others, total, comp, **STYLE["rev"])
    drv, off = [], []
    if rh and rp:
        a, b = (rp["rph"], rp["sh"]) if rp["rph_first"] else (rp["sh"], rp["rph"])
        if rp["same"]:
            drv += [a, b]
        else:
            drv.append(a)
            off.append(b)
        ride_sign = delta(rh, comp)
        if ride_sign * total < 0:  # ride-hailing va contra el total: queda como compensación
            off = drv + off
            drv = []
    elif rh:
        (drv if delta(rh, comp) * total > 0 else off).append(_item(rh, comp, ctx["kpis"], ctx=ctx))
    for x in om:
        drv.append(_item(x, comp, ctx["kpis"], ctx=ctx))
        evid.append(_ev(x, comp))
    for x in [x for x in others if x in oo]:
        off.append(_item(x, comp, ctx["kpis"], ctx=ctx))
        evid.append(_ev(x, comp))
    if rh and rh.get("note_m"):
        evid.insert(0, _ev(rh, comp))
    s = ""
    if len(drv) + len(off) >= 2:
        drv = [f"({ROMAN[i]}) {t}" for i, t in enumerate(drv)]
        off = [f"({ROMAN[len(drv) + i]}) {t}" for i, t in enumerate(off)]
    if drv:
        s = "Mainly driven by " + (", and ".join(drv) if len(drv) <= 2 else _join(drv)) + "."
    if off:
        s += (" This was partially offset by " if s else "Mainly driven by ") + (", and ".join(off) if len(off) <= 2 else _join(off)) + "."
    covered = sum(delta(n, comp) for n in ([rh] if rh else []) + om + oo)
    items_out = [(n["label"], delta(n, comp), 1) for n in ([rh] if rh else []) + om + oo]
    return {"text": f"{head}: {s.strip()}", "evidence": [e for e in evid if e], "check": (total, covered), "items": items_out}


def _ev(n, comp):
    if not n.get("note_m") and not (comp == "mom" and n.get("note_p")):
        return None
    out = {"line": n["label"], "delta": delta(n, comp), "row": n["row"], "note": n.get("note_m")}
    if comp == "mom" and n.get("note_p"):
        out["note_prev"] = n["note_p"]
    return out


# --------------------------------------------------------------------------
# API
# --------------------------------------------------------------------------

def open_wb(path):
    import openpyxl
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return openpyxl.load_workbook(path, data_only=True)


def prepare(wb, company, period, boxes=None, path=None):
    """Lee y arma el árbol una vez (lo caro); compose() se puede llamar muchas veces con distintos estilos."""
    ws = find_board_sheet(wb, company)
    data = read_board(ws, period, wb)
    roots = build_tree(data["rows"])
    nodes = {}
    for name, rx in (boxes or BOX_PATTERNS):
        n = _find(roots, rx)
        if n:
            nodes[name] = n
    return {"company": company, "period": period, "sheet": ws.title, "file": path, "data": data, "roots": roots,
            "nodes": nodes, "boxes": boxes or BOX_PATTERNS, "wb": wb, "ws": ws}


def compose(prep):
    data, nodes = prep["data"], prep["nodes"]
    ctx = {"kpis": data["kpis"], "days": data["days"], "box_keys": {n["key"] for n in nodes.values()},
           "prev_period": _prev(prep["period"]), "period": prep["period"], "flags": []}
    out = {"company": prep["company"], "period": prep["period"], "sheet": prep["sheet"], "file": prep["file"],
           "block": data["layout"]["block"], "boxes": []}
    for name in [b[0] for b in prep["boxes"]]:
        n = nodes.get(name)
        if not n:
            out["boxes"].append({"box": name, "missing": True})
            continue
        res = {"box": name, "actual": n["m"], "budget_value": n["b"], "prior": n["p"], "row": n["row"]}
        for comp in ("budget", "mom"):
            res[comp] = compose_box(name, n, comp, ctx)
        out["boxes"].append(res)

    def walk(ns):  # descuadres de budget dentro del árbol (renglón vs. suma de sus hijos)
        for x in ns:
            g = budget_gap(x)
            if g is not None and 0.05e6 <= abs(g) and abs(x["b"] or 0) > 0 and any(c["b"] is not None for c in x["children"]):
                ctx["flags"].append(f"{x['label']}: su budget no cuadra con la suma de sus renglones por {fm(g)} (revisar el archivo).")
            walk(x["children"])
    walk(prep["roots"])
    seen, flags = set(), []
    for f in ctx["flags"]:
        if f not in seen:
            seen.add(f)
            flags.append(f)
    out["flags"] = flags
    return out


def board_comments(path, company, period, boxes=None, style=True):
    if style:
        load_style()
    return compose(prepare(open_wb(path), company, period, boxes, path))


def latest_period(path, company=None):
    """Mes vigente del archivo = fecha del bloque 'Month Actual vs. Budget View'."""
    import openpyxl
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    ws = find_board_sheet(wb, company)
    maxc = min(ws.max_column, 200)
    for r in range(1, 10):
        for c in range(1, maxc + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and "month actual vs" in v.lower():
                for rr in range(r + 1, r + 12):
                    ym = _ym(ws.cell(rr, c).value)
                    if ym:
                        return ym
    raise RuntimeError("no pude leer el mes vigente del archivo")


def paste_text(res):
    """Solo los boxes, listos para pegar en la slide."""
    L = []
    for b in res["boxes"]:
        if b.get("missing"):
            continue
        L.append(f"{res['company']} – {b['box']}")
        L.append(b["budget"]["text"])
        L.append(b["mom"]["text"])
        L.append("")
    return "\n".join(L).strip() + "\n"


def render(res, show_evidence=True):
    L = [f"# {res['company']} — comentarios MFR {res['period']}",
         f"_Fuente: {res['sheet']} (serie mensual + columna Budget). Cifras de celda; el 'por qué' sale de las notas de celda del mes._", ""]
    for b in res["boxes"]:
        if b.get("missing"):
            L.append(f"## {b['box']}\n(no encontré el renglón en la hoja)\n")
            continue
        L.append(f"## {b['box']}")
        for comp in ("budget", "mom"):
            c = b[comp]
            L.append(c["text"])
            if c.get("check"):
                tot, cov = c["check"]
                if abs(tot - cov) >= 0.1e6:
                    L.append(f"<sub>✓ drivers {fm(cov)} de {fm(tot)} · resto {fm(tot - cov)} en partidas menores</sub>")
            L.append("")
    if res.get("flags"):
        L.append("## Revisar antes de pegar")
        L.extend(f"{i}. {f}" for i, f in enumerate(res["flags"], 1))
        L.append("")
    if show_evidence:
        L.append("## Soporte (notas de celda usadas)")
        seen = set()
        for b in res["boxes"]:
            if b.get("missing"):
                continue
            for comp in ("budget", "mom"):
                for e in b[comp]["evidence"]:
                    if e["row"] in seen:
                        continue
                    seen.add(e["row"])
                    if e.get("note"):
                        L.append(f"- **{e['line']}** ({res['period']}): {e['note'][:420]}")
                    if e.get("note_prev"):
                        L.append(f"- **{e['line']}** (mes anterior): {e['note_prev'][:300]}")
        L.append("")
    return "\n".join(L)


# ====================================================================================================
# contexto: notas de 3 meses + juntas
# ====================================================================================================

# renglón del Board View -> cómo se habla de él en juntas (español/inglés)
CTX_LINE_TERMS = {
    "ride hailing revenue": "ingresos rph supply hours precio por viaje ride hailing revenue",
    "advertising revenue": "publicidad advertising multimedia bull media cobranding ola media",
    "uber bonus": "bono uber bonus incentivos uber",
    "driver payments": "nomina conductores drivers pago a conductores sueldos salario minimo",
    "uber take rate": "uber take rate comision uber tr",
    "ev fuel cost": "energia electricidad cfe consumo de energia",
    "charging infra": "cargadores charging infraestructura hubs",
    "chaas": "chaas cargadores",
    "maintenance": "mantenimiento taller electromecanica correctivo preventivo siniestros",
    "washing": "lavado lavadores",
    "towing service": "grua gruas towing asistencia",
    "vehicle recovery services": "corralon corralones recuperacion",
    "non-operational fleet storage": "resguardo salazar pension",
    "vehicle leasing": "arrendamiento leasing unidades lto renta de autos",
    "insurance": "seguro seguros poliza siniestros deducible",
    "marketing": "marketing viniles branding facebook leads reclutamiento",
    "driving penalties & fines": "multas infracciones aeropuerto guardia nacional",
    "telematics": "telemetria samsara caska carstore gps",
    "security": "seguridad vigilancia",
    "vehicle administrative fees": "gestoria tenencias placas tramites",
    "payroll expense": "nomina headcount hc bono liquidaciones indemnizacion",
    "legal & professional fees": "honorarios legal sindicato imss despacho",
    "subscriptions": "licencias suscripciones software rankmi airbag",
    "others": "recargos iva multas sat comisiones",
    "insurance ": "sgmm seguro gastos medicos",
    "d&a": "depreciacion amortizacion",
    "taxes": "impuestos isr provision de impuestos",
    "cf & corp. g&a": "corporativo asignacion allocation cross functional",
}


def _date(s):
    try:
        return _dt.date.fromisoformat(str(s)[:10])
    except Exception:
        return None


def _window(period):
    """Juntas útiles para un cierre: del día 1 del mes al ~día 25 del mes subsiguiente (cuando se arma el MFR)."""
    y, m = int(period[:4]), int(period[5:7])
    start = _dt.date(y, m, 1)
    ny, nm = (y + (m + 1) // 12, (m + 1) % 12 + 1)
    return start, _dt.date(ny, nm, 25)


def _minus(period, n):
    y, m = int(period[:4]), int(period[5:7]) - n
    while m <= 0:
        y, m = y - 1, m + 12
    return f"{y:04d}-{m:02d}"


def notes_timeline(node, period, months=3):
    """Las notas de ese renglón en los últimos `months` meses calendario (incluye el actual): [(YYYY-MM, texto)]."""
    hist = node.get("notes_hist") or {}
    lo = _minus(period, months - 1)
    return [(k, hist[k]) for k in sorted(hist) if lo <= k <= period]


def meetings_for(mt, company, period, node, delta, k=2):
    """Fragmentos de juntas sobre ese renglón y ese cierre. 'match' = la junta dice la misma cifra que la
    variación (±10%) o que el monto del mes."""
    if mt is None or not len(mt):
        return []
    terms = CTX_LINE_TERMS.get(node["key"]) or _display(node["label"])
    start, end = _window(period)
    hits = mt.search(terms, k=40)
    out = []
    for h in hits:
        d = _date(h.get("date"))
        if not d or not (start <= d <= end):
            continue
        cos = h.get("companies") or []
        if cos and company not in cos:
            continue
        targets = [abs(delta)]  # solo la variación: el monto total del renglón empata por casualidad
        match = None
        for a in h.get("amounts") or []:
            for t in targets:
                if t >= 0.1e6 and abs(a - t) <= max(0.1 * t, 30_000):
                    match = (a, t)
                    break
            if match:
                break
        if h.get("similarity", 0) < 0.25 and not match:
            continue
        out.append(dict(h, match=match))
    out.sort(key=lambda h: (h["match"] is None, -h.get("similarity", 0)))
    return out[:k]


def build_context(res, prep, mt=None, min_abs=0.3e6):
    """Secciones de contexto para el reporte de un periodo.
    Devuelve {"lineas": [...], "juntas": [...], "pendientes": [...]}"""
    period, company = prep["period"], prep["company"]
    parent = {}

    def walk(ns, p=None):
        for x in ns:
            parent[id(x)] = p
            walk(x["children"], x)
    walk(prep["roots"])
    rows = [r for r in prep["data"]["rows"] if not r["children"]]
    lines = []
    for r in rows:
        p = parent.get(id(r))
        dm = delta(r, "mom")
        # sub-renglones que el budget no desglosa (p. ej. Wages dentro de Driver Payments): sin variación vs. budget
        no_b = r["b"] is None and p is not None and all(c["b"] is None for c in p["children"]) and p["b"] is not None
        db = 0.0 if no_b else delta(r, "budget")
        label = r["label"] if not p or p["key"] in ("opex", "other g&a", "revenue") else f"{r['label']} ({p['label']})"
        if max(abs(dm), abs(db)) < min_abs:
            continue
        lines.append({"label": label, "mom": dm, "budget": db, "notas": notes_timeline(r, period, 3), "node": r})
    lines.sort(key=lambda x: -max(abs(x["mom"]), abs(x["budget"])))
    juntas = []
    for x in lines[:14]:
        hs = meetings_for(mt, company, period, x["node"], x["mom"] if abs(x["mom"]) >= abs(x["budget"]) else x["budget"])
        if hs:
            juntas.append({"label": x["label"], "hits": hs})
    pend = []
    if mt is not None and len(mt):
        start, end = _window(period)
        for c in mt.chunks:
            d = _date(c.get("date"))
            if c.get("kind") == "pendiente" and d and start <= d <= end and (not c.get("companies") or company in c["companies"]):
                pend.append(c)
    return {"lineas": lines, "juntas": juntas, "pendientes": pend[:10]}


def render_context(ctx, period):
    L = []
    if ctx["lineas"]:
        L.append("## Qué pasó en el mes (tus notas de celda, últimos 3 meses por renglón)")
        for x in ctx["lineas"][:16]:
            head = f"- **{x['label']}** (MoM {fm(x['mom'])} · vs. Budget {fm(x['budget'])})"
            if not x["notas"]:
                L.append(head + ": _sin notas en el Excel_")
                continue
            L.append(head)
            for ym, t in reversed(x["notas"]):
                tag = "este mes" if ym == period else ym
                L.append(f"    - _{tag}_: {t[:350]}")
        L.append("")
    if ctx["juntas"]:
        L.append("## Lo que se dijo en juntas (Granola) sobre estos renglones")
        for j in ctx["juntas"]:
            for h in j["hits"]:
                ok = f" — **✔ cifra dicha {fm(h['match'][0])} ≈ Excel {fm(h['match'][1])}**" if h.get("match") else ""
                link = f" ([nota]({h['url']}))" if h.get("url") else ""
                L.append(f"- **{j['label']}** · {h['meeting']} ({h['date']}){link}: “{h['text'][:380]}”{ok}")
        L.append("")
    if ctx.get("dynamics_md"):
        L += ctx["dynamics_md"] + [""]
    if ctx["pendientes"]:
        L.append("## Pendientes de juntas de este cierre")
        for c in ctx["pendientes"]:
            L.append(f"- {c['text'][:260]} _({c['meeting']}, {c['date']})_")
        L.append("")
    return "\n".join(L)


# ====================================================================================================
# tus boxes publicados en los decks (oro)
# ====================================================================================================

GOLD_MONTHS = {"january": 1, "enero": 1, "february": 2, "febrero": 2, "march": 3, "marzo": 3, "april": 4, "abril": 4,
          "may": 5, "mayo": 5, "june": 6, "junio": 6, "july": 7, "julio": 7, "august": 8, "agosto": 8,
          "september": 9, "sept": 9, "septiembre": 9, "october": 10, "oct": 10, "octubre": 10, "november": 11,
          "nov": 11, "noviembre": 11, "december": 12, "dic": 12, "diciembre": 12}
GOLD_COMPANIES = ["DAE", "LTO", "EVF", "VCN"]
GOLD_BOX_TITLES = r"^(revenues?|gross profit|gp|ebitda|net income|net operating revenue|financial margin|adjusted financial margin|opex|sg&a)$"
GOLD_BOX_NORMAL = {"gp": "Gross Profit", "revenue": "Revenues"}
GOLD_SKIP = r"copia|backup|deltas|pre-guion|\bmkt\b|q[1-4]\W|@corp|performance|portfolio|hl & ll|mbr|vcn[ _-]|sent vrp|w glcomments|previous"


def gold_period_of(path):
    p = path.replace("\\", "/").lower()
    y = re.search(r"/(20\d\d)/", p)
    name = os.path.basename(p)
    for k, v in GOLD_MONTHS.items():
        if re.search(rf"\b{k}\b", name):
            return f"{y.group(1)}-{v:02d}" if y else None
    folder = os.path.basename(os.path.dirname(p))
    m = re.match(r"(\d\d)\s", folder)
    return f"{y.group(1)}-{int(m.group(1)):02d}" if (y and m) else None


def _last_closed():
    import datetime as _d
    t = _d.date.today()
    y, m = (t.year, t.month - 1) if t.month > 1 else (t.year - 1, 12)
    return f"{y:04d}-{m:02d}"


def pick_decks(root):
    """Un deck por mes: el más reciente que no sea copia/respaldo/sección parcial."""
    best = {}
    for dp, dn, fn in os.walk(root):
        if "claude outputs" in dp.lower():
            continue
        for f in fn:
            fl = f.lower()
            if not fl.endswith(".pptx") or not re.search(r"financial|^vemo - .*monthly report", fl) or re.search(GOLD_SKIP, fl):
                continue
            path = os.path.join(dp, f)
            per = gold_period_of(path)
            if not per or per > _last_closed():  # un deck con nombre de mes futuro es una copia de otro año
                continue
            try:
                st = os.stat(path)
            except OSError:
                continue
            if st.st_size < (3e6 if per < "2024-01" else 20e6):  # los decks de 2023 pesaban menos
                continue
            score = (("presented" in fl) * 2 + ("board" in fl) + ("audited" not in fl)
                     + bool(re.search(r"\bsent\b|final|\bvf\b|\(fv\)", fl)) - bool(re.search(r"draft|hipo|_v1|v1\b", fl)), st.st_mtime)
            if per not in best or score > best[per][0]:
                best[per] = (score, path)
    return {k: v[1] for k, v in sorted(best.items())}


def _gold_shapes(shapes, out, off=(0, 0)):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    for s in shapes:
        try:
            if s.shape_type == MSO_SHAPE_TYPE.GROUP:
                _gold_shapes(s.shapes, out)
                continue
        except Exception:
            pass
        if getattr(s, "has_text_frame", False) and s.has_text_frame:
            t = s.text_frame.text.strip()
            if t and s.left is not None and s.top is not None:
                out.append({"text": t, "x": s.left, "y": s.top, "w": s.width or 0, "h": s.height or 0})


def gold_split_comp(text):
    """'Budget (MXN ...): ...\\nMoM (MXN ...): ...' -> {'budget': ..., 'mom': ...}"""
    out = {}
    parts = re.split(r"(?=(?:^|\n)\s*(?:Budget|MoM|vs\.? ?Budget)\s*[:(])", text)
    for p in parts:
        p = p.strip()
        if re.match(r"(vs\.? ?)?budget", p, re.I):
            out["budget"] = p
        elif re.match(r"mom", p, re.I):
            out["mom"] = p
    return out


def _hl_items(sh, co, period, path, slide):
    """Slide '<CO> – Highlights & Lowlights': cada párrafo 'Tema: texto' bajo su columna."""
    heads = {k: next((x for x in sh if x["text"].strip().lower() == k), None) for k in ("highlights", "lowlights")}
    if not heads["highlights"] or not heads["lowlights"]:
        return []
    top = min(heads["highlights"]["y"], heads["lowlights"]["y"])
    out = []
    for x in sorted(sh, key=lambda x: (x["x"], x["y"])):
        if x["y"] <= top or x in heads.values() or len(x["text"]) < 25:
            continue
        cx = x["x"] + x["w"] / 2
        kind = min(heads, key=lambda k: abs(cx - (heads[k]["x"] + heads[k]["w"] / 2)))
        cur = None
        for par in [p.strip() for p in x["text"].split("\n") if p.strip()]:
            m = re.match(r"^([^:]{3,70}):\s*(.+)$", par)
            if m or cur is None:
                cur = {"company": co, "period": period, "deck": os.path.basename(path), "slide": slide, "kind": kind,
                       "tema": (m.group(1) if m else "").strip(" –-"), "text": (m.group(2) if m else par).strip()}
                out.append(cur)
            else:
                cur["text"] += " " + par
    return out


def _co_from_label(t):
    t = _plain(t)
    if re.search(r"\bdae\b", t) and "colombia" not in t:
        return "DAE"
    if re.search(r"\blto\b", t):
        return "LTO"
    if re.search(r"ev ?fleets|\bevf\b", t):
        return "EVF"
    if re.search(r"charging network|\bvcn\b", t):
        return "VCN"
    return None


def _hl_items_multi(sh, period, path, slide):
    heads = {k: next((x for x in sh if x["text"].strip().lower() == k), None) for k in ("highlights", "lowlights")}
    if not heads["highlights"] or not heads["lowlights"]:
        return []
    top = min(heads["highlights"]["y"], heads["lowlights"]["y"])
    left = heads["highlights"]["x"] + 0.25 * heads["highlights"]["w"]
    labels = sorted([x for x in sh if x["y"] > top and x["x"] < left and x["w"] < 0.25 * heads["highlights"]["w"]
                     and _co_from_label(x["text"])], key=lambda x: x["y"])
    out = []
    for x in sh:
        if x["y"] <= top or x in heads.values() or x in labels or len(x["text"]) < 25:
            continue
        lab = max((lb for lb in labels if lb["y"] <= x["y"] + 0.5 * max(lb["h"], 1)), key=lambda lb: lb["y"], default=None)
        if not lab:
            continue
        cx = x["x"] + x["w"] / 2
        kind = min(heads, key=lambda k: abs(cx - (heads[k]["x"] + heads[k]["w"] / 2)))
        for par in [p.strip() for p in re.split(r"\n+", x["text"]) if len(p.strip()) > 20]:
            m = re.match(r"^([^:]{3,70}):\s*(.+)$", par)
            out.append({"company": _co_from_label(lab["text"]), "period": period, "deck": os.path.basename(path), "slide": slide,
                        "kind": kind, "tema": (m.group(1) if m else "").strip(" –-\u200b"), "text": (m.group(2) if m else par).strip()})
    return out


def gold_extract(path, period):
    from pptx import Presentation
    prs = Presentation(path)
    recs = []
    for i, slide in enumerate(prs.slides, 1):
        sh = []
        _gold_shapes(slide.shapes, sh)
        hl = next((s for s in sh if re.search(r"\b(DAE|LTO|EVF|VCN)\b\s*[–-]?\s*Highlights", s["text"])), None)
        if hl:
            recs += _hl_items(sh, re.search(r"\b(DAE|LTO|EVF|VCN)\b", hl["text"]).group(1), period, path, i)
            continue
        hl23 = next((s for s in sh if re.match(r"^\s*Highlights & Lowlights\s*[–-]", s["text"])), None)
        if hl23:  # 2023: una slide con varias empresas, la empresa en una etiqueta a la izquierda
            recs += _hl_items_multi(sh, period, path, i)
            continue
        old_pl = next((s["text"] for s in sh if re.match(r"^\s*(DAE|LTO|EV Fleets|Charging Network)\s*:\s*P&L Summary", s["text"])), None)
        if old_pl:  # 2023: comentarios 'Vs. Runway' (formato anterior; solo contexto, no para el estilo)
            co = _co_from_label(old_pl.split(":")[0])
            for c in sh:
                if re.search(r"runway|prior month|vs\.? ?(previous|last) month", c["text"], re.I) and len(c["text"]) > 60:
                    recs.append({"company": co, "period": period, "deck": os.path.basename(path), "slide": i,
                                 "kind": "comentario_pl", "tema": "P&L Summary", "text": c["text"].strip()})
            continue
        title = next((s["text"] for s in sh if re.search(r"\b(DAE|LTO|EVF|VCN)\b(?! \(Colombia\))[^–\-:\n]{0,12}[–\-:]\s*(P&L Summary|Operational & Financial Results)", s["text"])), None)
        if not title:
            continue
        co = re.search(r"\b(DAE|LTO|EVF|VCN)\b", title).group(1)
        titles = [s for s in sh if re.match(GOLD_BOX_TITLES, s["text"].strip().lower())]
        comments = [s for s in sh if re.search(r"(^|\n)\s*(Budget|MoM)\s*\(", s["text"])]
        if not comments and any(x["text"].strip().lower() == "lowlights" for x in sh):  # 2024: HL operativos
            recs += _hl_items(sh, co, period, path, i)
            continue
        for c in comments:
            # el título del box está justo arriba del comentario y se traslapa en horizontal
            def above(t):
                overlap = min(t["x"] + t["w"], c["x"] + c["w"]) - max(t["x"], c["x"])
                return overlap > 0 and t["y"] <= c["y"] + 0.15 * c["h"]
            cands = [t for t in titles if above(t)]
            t = max(cands, key=lambda t: t["y"]) if cands else None
            box = t["text"].strip() if t else None
            if box:
                box = GOLD_BOX_NORMAL.get(box.lower(), box)
            for comp, txt in gold_split_comp(c["text"]).items():
                m = re.match(r"[^(]*\(\s*(?:ca\.?\s*)?MXN\s*([+\-‑−]?)\s*\$?\s*([\d.,]+)\s*m", txt)
                amt = None
                if m:
                    amt = float(m.group(2).replace(",", "")) * (-1 if m.group(1) in "-‑−" and m.group(1) else 1)
                recs.append({"company": co, "period": period, "deck": os.path.basename(path), "slide": i,
                             "box": box, "comp": comp, "amount_m": amt, "text": txt})
    return recs


def gold_update(root, out, verbose=True):
    decks = pick_decks(root)
    allr = []
    if os.path.exists(out):
        try:
            allr = json.load(open(out, encoding="utf-8"))
        except Exception:
            allr = []
    done = {(r["period"], r["deck"], r.get("deck_mtime")) for r in allr}
    idx_path = out + ".idx.json"  # decks ya leídos (aunque no traigan boxes), para no releerlos
    try:
        idx = json.load(open(idx_path, encoding="utf-8"))
    except Exception:
        idx = {}
    done |= {(v["period"], k, v["mtime"]) for k, v in idx.items()}
    hl_path = os.path.join(os.path.dirname(out), "mfr_highlights.json")
    try:
        hls = json.load(open(hl_path, encoding="utf-8"))
    except Exception:
        hls = None  # primera vez: releo todos los decks para sacar Highlights & Lowlights
    if hls is None:
        done, hls = set(), []
    lo, hi = os.environ.get("MFR_DESDE", "2023-01"), os.environ.get("MFR_HASTA", "9999-12")
    for per, path in decks.items():
        if not (lo <= per <= hi):
            continue
        mt = int(os.stat(path).st_mtime)
        if (per, os.path.basename(path), mt) in done and os.environ.get("MFR_REHACER") != "1":
            continue
        try:
            recs = gold_extract(path, per)
        except Exception as e:
            print("!!", per, path, e)
            continue
        hrecs = [r for r in recs if "kind" in r]
        recs = [r for r in recs if "kind" not in r]
        hls = [r for r in hls if r["period"] != per] + hrecs
        json.dump(hls, open(hl_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        for r in recs:
            r["deck_mtime"] = mt
        idx[os.path.basename(path)] = {"period": per, "mtime": mt, "n": len(recs)}
        json.dump(idx, open(idx_path, "w", encoding="utf-8"), ensure_ascii=False)
        allr = [r for r in allr if r["period"] != per] + recs
        json.dump(allr, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        if verbose:
            print(f"  deck {per}: {os.path.basename(path)} -> {len(recs)} comentarios tuyos, {len(hrecs)} highlights/lowlights", flush=True)
    if not os.path.exists(hl_path):
        json.dump(hls, open(hl_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    keep = [r for r in allr if r["period"] <= _last_closed()]
    if len(keep) != len(allr):
        json.dump(keep, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return keep


# ====================================================================================================
# aprendizaje de tu estilo + calificación
# ====================================================================================================

LEARN_MINUS_RX = "[\\-‑−]"

LEARN_ALIASES = [  # frase tuya -> renglón del P&L (clave normalizada de la hoja)
    (r"uber (tr|take rate)", "uber take rate"),
    (r"driver payments?|driver payroll|payroll costs?|labor costs?", "driver payments"),
    (r"maintenance( & towing)?", "maintenance"),
    (r"washing", "washing"),
    (r"towing", "towing service"),
    (r"marketing", "marketing"),
    (r"vehicle leasing|lto vehicle leasing", "vehicle leasing"),
    (r"driving (penalties|fines)|penalties", "driving penalties & fines"),
    (r"telematics", "telematics"),
    (r"vehicle administrative fees|administrative fees", "vehicle administrative fees"),
    (r"charging infra", "charging infra"),
    (r"ev fuel", "ev fuel cost"),
    (r"advertising", "advertising revenue"),
    (r"ride.hailing", "ride hailing revenue"),
    (r"\bopex\b", "opex"),
    (r"^gp$|gross profit|\bgp\b", "gross profit"),
    (r"\b(dl|dbl) (expenses|g&a)|direct (line )?sg&a|direct business line", "direct line sg&a"),
    (r"cf (&|and) corp", "cf & corp. g&a"),
    (r"payroll expenses?", "payroll expense"),
    (r"subscriptions?", "subscriptions"),
    (r"legal", "legal & professional fees"),
    (r"\bebitda\b", "ebitda"),
    (r"\bd&a\b", "d&a"),
    (r"\btaxes\b", "taxes"),
    (r"interest expense", "interest expense"),
    (r"\brevenues?\b", "revenue"),
    (r"insurance", "insurance"),
    (r"security", "security"),
]


def _plain(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def classify_box(text):
    t = text
    if re.search(r"(higher|lower) EBITDA \(", t) or re.search(r"\bD&A\b.*\btaxes\b", t, re.I | re.S):
        return "Net Income"
    if re.search(r"\b(higher|lower) GP\b|\bGP \(|\bDL\b|\bDBL\b|Direct SG&A|CF (&|and) Corp", t):
        return "EBITDA"
    if re.search(r"\bRPH\b|revenue per hour|\bSH\b|price per ride|\bPPR\b", t, re.I):
        return "Revenues"
    if re.search(r"OPEX|Uber T|driver pay|payroll cost|maintenance|Revenues? \(", t, re.I):
        return "Gross Profit"
    return None


def gold_records(path, company):
    """Tus boxes de la empresa, con el box corregido por contenido (el título del deck a veces no se lee bien)."""
    with open(path, encoding="utf-8") as fh:
        recs = [r for r in json.load(fh) if r["company"] == company]
    out, i = [], 0
    while i < len(recs):  # budget y mom del mismo recuadro van juntos
        r = recs[i]
        pair = [r]
        if i + 1 < len(recs) and recs[i + 1]["slide"] == r["slide"] and recs[i + 1]["comp"] != r["comp"] \
                and recs[i + 1]["period"] == r["period"]:
            pair.append(recs[i + 1])
        box = classify_box(" ".join(p["text"] for p in pair)) or r.get("box")
        for p in pair:
            q = dict(p)
            q["box"] = box
            out.append(q)
        i += len(pair)
    return out


# --------------------------------------------------------------------------
# parseo de tus comentarios
# --------------------------------------------------------------------------

AMT_IN_PAREN = re.compile(r"\(\s*(?:ca\.?\s*)?(?:MXN\s*)?(\+|" + LEARN_MINUS_RX + r")?\s*(?:ca\.?\s*)?\$\s*(\d+(?:\.\d+)?)\.?\s*[mM]?\b")


def learn_line_key(phrase, box):
    p = _plain(phrase)
    if box == "Gross Profit" and re.search(r"payroll", p):
        return "driver payments"
    for rx, key in LEARN_ALIASES:
        if re.search(rx, p):
            return key
    return None


def learn_mentions(text, box):
    """[(clave, frase, monto_m, razon, pos)] de cada driver con monto en tu comentario."""
    body = text.split(":", 1)[1] if ":" in text else text
    body = re.sub(r"\bCF and Corp", "CF & Corp", body)
    out = []
    ms = list(AMT_IN_PAREN.finditer(body))
    for i, m in enumerate(ms):
        before = body[:m.start()]
        seg = re.split(r"[;:]|\.\s+(?=[A-Z][a-z])|\.$|\(\s*[ivx]+\s*\)|\(\s*[a-h]\s*\)|,|\band\b", before)[-1]
        phrase = re.sub(r"^\s*(mainly|primarily|driven by|partially offset by|offset by|this was|by|due to|explained by|"
                        r"higher|lower|slightly|a|the|other)\b\s*", "", seg.strip(), flags=re.I)
        phrase = re.sub(r"^\s*(higher|lower|slightly|a|the)\s+", "", phrase, flags=re.I).strip(" ,")
        sign = -1 if (m.group(1) and m.group(1) != "+") else 1
        amt = sign * float(m.group(2))
        end = ms[i + 1].start() if i + 1 < len(ms) else len(body)
        after = body[m.end():end]
        after = after.split(")", 1)[1] if ")" in after[:40] else after
        reason = re.split(r";|\(\s*[ivx]+\s*\)|\(\s*[a-h]\s*\)|\.\s|This was|partially offset", after)[0].strip(" ,.")
        key = learn_line_key(phrase, box)
        out.append({"key": key, "phrase": phrase, "amt": amt, "reason": reason, "paren": body[m.start():m.end() + 40]})
    return out


def kpi_pairs(text):
    """'(1,664 vs. 1,415)' / '(13.8% vs. 9.2%)' / '($0.97 vs. $1.32)' -> [(x, y, tipo, frase_previa)]."""
    out = []
    for m in re.finditer(r"\(\s*\$?([\d,]+(?:\.\d+)?)\s*(%|k|K)?\s*vs\.?\s*\$?([\d,]+(?:\.\d+)?)\s*(%|k|K)?\s*\)", text):
        x, y = float(m.group(1).replace(",", "")), float(m.group(3).replace(",", ""))
        kind = "pct" if m.group(2) == "%" else ("k" if (m.group(2) or "").lower() == "k" else
                                                ("money" if "$" in text[m.start():m.start() + 3] else "int"))
        pre = text[:m.start()]
        phrase = re.split(r"[;,(]|\)\s", pre)[-1].strip()
        out.append((x, y, kind, phrase, m.start()))
    return out


# --------------------------------------------------------------------------
# calificación de un box generado contra el tuyo
# --------------------------------------------------------------------------

def _eng_keys(items):
    keys = []
    for lab, d, lvl in items or []:
        k = _plain(lab)
        if k.startswith("other "):
            continue
        keys.append(k)
    return keys


LEARN_KPI_TOKENS = [("RPH", r"\bRPH\b|revenue per hour"), ("PPR", r"\bPPR\b|price per ride"), ("SH", r"\bSH\b|supply hours")]


def learn_keyset(text, box):
    """Lo que dice un box: renglones con monto + KPIs del ride-hailing (se aplica igual a tu texto y al de la IA)."""
    ks = {m["key"] for m in learn_mentions(text, box) if m["key"]}
    if box == "Revenues":
        ks |= {name for name, rx in LEARN_KPI_TOKENS if re.search(rx, text, re.I)}
        ks.discard("ride hailing revenue")
        ks.discard("revenue")
    return ks


def learn_score(gold_text, gen, box):
    g = learn_keyset(gold_text, box)
    e = learn_keyset(gen["text"], box)
    inline_g = bool(re.search(r"in line", gold_text, re.I)) and not g
    inline_e = gen["text"].rstrip().endswith("In line.")
    if inline_g or inline_e:
        return 1.0 if inline_g == inline_e else 0.0, g, e
    if not g and not e:
        return 1.0, g, e
    if not g or not e:
        return 0.0, g, e
    tp = len(g & e)
    p, r = tp / len(e), tp / len(g)
    return (2 * p * r / (p + r) if tp else 0.0), g, e


# --------------------------------------------------------------------------
# aprendizaje
# --------------------------------------------------------------------------

def learn_vocab(gold_all, preps, company):
    """display, no_expand, expand_word, recurring, ni_mom_inline, kpi_learned, slightly_pct."""
    disp = defaultdict(Counter)
    exp_yes, exp_no = Counter(), Counter()
    words = {"budget": Counter(), "mom": Counter()}
    reasons = defaultdict(Counter)
    inline = Counter()
    kpi = defaultdict(Counter)
    kpi_phr = {}
    slight = []
    for r in gold_all:
        box, comp, text = r["box"], r["comp"], r["text"]
        if not box:
            continue
        y, mth = int(r["period"][:4]), int(r["period"][5:7])
        w = max(1, (y - 2025) * 12 + mth - 6)  # tu estilo reciente pesa más (lineal en el tiempo)
        for m in learn_mentions(text, box):
            if not m["key"]:
                continue
            ph = m["phrase"].strip()
            for _ in range(3):
                ph = re.sub(r"^\s*(partially offset by|offset by|slightly offset by|partially|offset|by|and|of|higher|lower|"
                            r"slightly|a|the|other|increased|reduced)\b\s*", "", ph, flags=re.I)
            if 2 <= len(ph) <= 40 and not re.search(r"\d", ph) and (not company or r["company"] == company) \
                    and learn_line_key(ph, box) == m["key"]:
                disp[m["key"]][ph] += w
            follows = re.match(r"\s*,?\s*(explained by|primarily due to|primarily explained by|mainly from|due to)\s*:?\s*\(a\)", m["paren"][m["paren"].find(")") + 1:] if ")" in m["paren"] else "")
            if not company or r["company"] == company:
                if follows:
                    exp_yes[m["key"]] += w
                else:
                    exp_no[m["key"]] += w
            rs = m["reason"]
            content = re.sub(r"^(from|due to|driven by|on account of)\b", "", rs or "").strip(" :,.")
            if rs and 6 <= len(rs) <= 60 and not re.search(r"\d", rs) and re.match(r"(from|due to|driven by|on account of)\b", rs) \
                    and len(content.split()) >= 2 and not rs.rstrip().endswith(":") \
                    and m["key"] not in ("opex", "revenue", "direct line sg&a", "gross profit", "ebitda", "cf & corp. g&a"):
                dirn = "up" if re.search(r"higher|increase", m["phrase"] + " " + text[:0], re.I) or m["amt"] < 0 else "down"
                reasons[f"{m['key']}|*|{dirn}"][rs] += 1
        for kw in ("explained by:", "primarily due to:", "primarily explained by"):
            words[comp][kw] += len(re.findall(re.escape(kw), text)) * w
        if box == "Net Income" and comp == "mom":
            inline["yes" if re.search(r"in line", text, re.I) else "no"] += w
        for sm in re.finditer(r"slightly (higher|lower) SH\s*([+\-‑]?[\d.]+)%", text):
            slight.append(abs(float(sm.group(2).replace("‑", "-"))))
        # KPI por valor: "(x vs. y)" junto a un driver -> qué renglón de KPI de la hoja es
        if company and r["company"] == company and r["period"] in preps:
            kp = preps[r["period"]]["data"]["kpis"]
            ms = learn_mentions(text, box)
            for x, y, kind, phrase, pos in kpi_pairs(text.split(":", 1)[1] if ":" in text else text):
                owner = None
                body = text.split(":", 1)[1] if ":" in text else text
                prev = [m for m in ms if body.find(m["paren"][:14]) != -1 and body.find(m["paren"][:14]) < pos]
                if prev:
                    owner = prev[-1]["key"]
                if not owner or owner in ("revenue", "ride hailing revenue"):
                    continue
                for kk, rec in kp.items():
                    a, ref = rec["m"], (rec["b"] if comp == "budget" else rec["p"])
                    if a is None or ref is None:
                        continue
                    sc = 100.0 if kind == "pct" else 1.0
                    if abs(a * sc - x) <= max(0.051, 0.006 * abs(x)) and abs(ref * sc - y) <= max(0.051, 0.006 * abs(y)) \
                            or (kind == "k" and abs(a / 1000 - x) <= 0.06 and abs(ref / 1000 - y) <= 0.06):
                        kpi[owner][(kk, kind)] += 1
                        if phrase and len(phrase) <= 50 and not re.search(r"\d", phrase):
                            kpi_phr[(owner, kk, "up" if x > y else "down")] = re.sub(r"^\s*(and|,)\s*", "", phrase)
    style = {}
    style["display"] = {}
    for key, c in disp.items():
        ph, n = c.most_common(1)[0]
        if n >= 2 and key not in ("revenue", "ebitda"):
            style["display"][key] = ph
    style["always_expand"] = [k for k in exp_yes if exp_yes[k] >= 10 and exp_yes[k] >= 2 * exp_no[k]]
    style["no_expand"] = [k for k in set(exp_no) | set(exp_yes) if exp_no[k] >= 10 and exp_yes[k] <= 0.1 * exp_no[k] and k in (
        "cf & corp. g&a", "d&a", "taxes", "interest expense", "gross profit", "revenue")]
    style["expand_word"] = {c: (words[c].most_common(1)[0][0] if words[c] else "explained by:") for c in ("budget", "mom")}
    style["expand_word"] = {c: (w if w.endswith(":") else w + ":") for c, w in style["expand_word"].items()}
    style["recurring"] = {k: c.most_common(1)[0][0] for k, c in reasons.items() if c.most_common(1)[0][1] >= 3}
    style["ni_mom_inline"] = inline["yes"] > inline["no"]
    style["slightly_pct"] = round(max(slight) + 0.05, 2) if slight else 0.0
    kl = {}
    for owner, c in kpi.items():
        (kk, kind), n = c.most_common(1)[0]
        if n < 2:
            continue
        up = kpi_phr.get((owner, kk, "up"))
        down = kpi_phr.get((owner, kk, "down"))
        if not up and not down:
            continue
        up = up or _learn_flip(down)
        down = down or _learn_flip(up)
        kl[owner] = [{"kpi": kk, "fmt": kind, "up": up, "down": down, "n": n}]
    style["kpi_learned"] = kl
    return style


def _learn_flip(p):
    if not p:
        return p
    sw = {"higher": "lower", "lower": "higher", "additional": "fewer", "fewer": "additional", "increase": "decrease",
          "decrease": "increase", "more": "fewer"}
    return re.sub(r"\b(" + "|".join(sw) + r")\b", lambda m: sw[m.group(1)], p)


def _assign_boxes_by_header(gold, preps):
    """El encabezado (MXN ±$x.xm) de tu box identifica el renglón: se compara contra las variaciones del archivo."""
    by = defaultdict(list)
    for g in gold:
        by[(g["period"], g["slide"], g["deck"])].append(g)
    for (per, _s, _d), recs in by.items():
        nodes = preps[per]["nodes"]
        # pares budget/mom en orden
        i = 0
        while i < len(recs):
            pair = [recs[i]]
            if i + 1 < len(recs) and recs[i + 1]["comp"] != recs[i]["comp"]:
                pair.append(recs[i + 1])
            best, err_best = None, 1e9
            errs = {}
            for name, n in nodes.items():
                err, cnt = 0.0, 0
                for g in pair:
                    if g.get("amount_m") is None:
                        continue
                    d = delta(n, g["comp"]) / 1e6
                    err += abs(d - g["amount_m"])
                    cnt += 1
                if cnt:
                    errs[name] = err / cnt
                if cnt and err / cnt < err_best:
                    best, err_best = name, err / cnt
            content = classify_box(" ".join(g["text"] for g in pair))
            if content in errs and errs[content] <= 0.25:  # empate de encabezados (p. ej. GP = EBITDA): manda el contenido
                best, err_best = content, errs[content]
            for g in pair:
                g["head_ok"] = bool(best and err_best <= 0.25)
            if best and err_best <= 0.25:
                for g in pair:
                    g["box"] = best
            i += len(pair)


LEARN_GRID = [
    ("top", "max_n", [2, 3, 4, 5]),
    ("top", "min_abs", [0.05e6, 0.2e6, 0.4e6, 0.7e6, 1.0e6]),
    ("sub", "max_n", [2, 3, 4, 5, 7]),
    ("sub", "min_abs", [0.1e6, 0.2e6, 0.3e6, 0.5e6]),
    ("sub", "cover", [0.6, 0.75, 0.85, 0.95]),
    ("sub", "expand_min", [0.3e6, 0.5e6, 1.0e6, 2.0e6]),
    ("offs", "rel", [0.03, 0.1, 0.2, 0.35]),
    ("offs", "max_n", [0, 1, 2, 3]),
    ("resid", "on", [True, False]),
    ("rev", "min_abs", [0.1e6, 0.3e6, 0.5e6, 1.0e6]),
    ("offs", "min_abs", [0.1e6, 0.2e6, 0.4e6]),
    ("ni_mom_inline", None, [True, False]),
    ("slightly_pct", None, [0.0, 1.0, 2.0]),
]


def learn_evaluate(style, preps, gold):
    set_style(style)
    rows = []
    cache = {}
    for g in gold:
        per, box, comp = g["period"], g["box"], g["comp"]
        if per not in preps or not box:
            continue
        if per not in cache:
            cache[per] = compose(preps[per])
        b = next((x for x in cache[per]["boxes"] if x["box"] == box and not x.get("missing")), None)
        if not b:
            continue
        gen = b[comp]
        s, gk, ek = learn_score(g["text"], gen, box)
        head_gen = float(re.search(r"\(MXN ([+‑\-])\$([\d.]+)m\)", gen["text"]).group(2)) * \
            (-1 if re.search(r"\(MXN [‑\-]", gen["text"]) else 1) if re.search(r"\(MXN [+‑\-]\$", gen["text"]) else None
        rows.append({"period": per, "box": box, "comp": comp, "learn_score": s, "gold_keys": sorted(gk), "gen_keys": sorted(ek),
                     "head_gold": g.get("amount_m"), "head_gen": head_gen, "gold": g["text"], "gen": gen["text"]})
    return (sum(r["learn_score"] for r in rows) / len(rows) if rows else 0.0), rows


def learn_tune(base, preps, gold, passes=2):
    best = copy.deepcopy(base)
    best_s, _ = learn_evaluate(best, preps, gold)
    for _ in range(passes):
        improved = False
        for sec, k, vals in LEARN_GRID:
            for v in vals:
                cand = copy.deepcopy(best)
                if k is None:
                    cand[sec] = v
                else:
                    cand[sec][k] = v
                s, _ = learn_evaluate(cand, preps, gold)
                if s > best_s + 1e-9:
                    best, best_s, improved = cand, s, True
        if not improved:
            break
    return best, best_s


LEARN_BOX_GRID = [
    ("top", "max_n", [2, 3, 4, 5]),
    ("top", "min_abs", [0.05e6, 0.2e6, 0.4e6, 0.7e6, 1.0e6]),
    ("sub", "max_n", [2, 3, 4, 5, 7]),
    ("sub", "min_abs", [0.1e6, 0.2e6, 0.3e6, 0.5e6]),
    ("sub", "expand_min", [0.3e6, 0.5e6, 1.0e6, 2.0e6]),
    ("offs", "rel", [0.03, 0.1, 0.2, 0.35]),
    ("offs", "max_n", [0, 1, 2, 3]),
]


def tune_per_box(style, preps, gold, passes=1):
    """Ajuste fino por box: si escribes distinto el GP que el EBITDA, cada uno aprende sus propios límites."""
    best = copy.deepcopy(style)
    best_s, _ = learn_evaluate(best, preps, gold)
    for box in ("Gross Profit", "EBITDA", "Net Income"):
        for _ in range(passes):
            for sec, k, vals in LEARN_BOX_GRID:
                for v in vals:
                    cand = copy.deepcopy(best)
                    pb = cand.setdefault("per_box", {}).setdefault(box, {})
                    pb.setdefault(sec, {})[k] = v
                    s, _ = learn_evaluate(cand, preps, gold)
                    if s > best_s + 1e-9:
                        best, best_s = cand, s
    return best, best_s


def learn_run(company, fy_path, gold_path, periods=None, out_path=None, report_path=None, verbose=True):
    t0 = time.time()
    with open(gold_path, encoding="utf-8") as fh:
        gold_all_raw = json.load(fh)
    gold_all = []
    for co in sorted({r["company"] for r in gold_all_raw}):
        gold_all += gold_records(gold_path, co)
    gold = [g for g in gold_all if g["company"] == company and g["box"]]
    pers = sorted({g["period"] for g in gold if (not periods or g["period"] in periods)})
    wb = open_wb(fy_path)
    preps = {}
    for per in pers:
        try:
            preps[per] = prepare(wb, company, per, path=fy_path)
        except Exception as e:
            if verbose and "no tiene la columna" not in str(e):  # meses de otros años: normal, no es error
                print(f"  {per}: sin datos en el archivo ({e})")
    for p in list(preps):  # meses sin presupuesto o sin cierre en el archivo no sirven para calificar
        nodes = preps[p]["nodes"].values()
        if all(n["b"] is None for n in nodes) or all(not n["m"] for n in nodes):
            preps.pop(p)
    gold = [g for g in gold if g["period"] in preps]
    _assign_boxes_by_header(gold, preps)
    # un mes cuyo deck todavía trae los boxes del mes anterior (deck en proceso) no sirve de oro: se descarta
    bad = []
    for p in sorted(preps):
        prev = _prev(p)
        if prev not in preps:
            continue
        stale = total = 0
        for g in gold:
            if g["period"] != p or g.get("amount_m") is None or g["comp"] != "budget":
                continue
            total += 1
            cur = min(abs(delta(n, "budget") / 1e6 - g["amount_m"]) for n in preps[p]["nodes"].values())
            old = min(abs(delta(n, "budget") / 1e6 - g["amount_m"]) for n in preps[prev]["nodes"].values())
            stale += old < 0.06 and cur > 0.25
        if total and stale > 0.5 * total:
            bad.append(p)
    # comentario por comentario: si su encabezado es el del mes anterior (box aún no actualizado), se descarta
    drop = set()
    for i, g in enumerate(gold):
        p, prev = g["period"], _prev(g["period"])
        if p in bad or prev not in preps or g.get("amount_m") is None:
            continue
        cur = min(abs(delta(n, g["comp"]) / 1e6 - g["amount_m"]) for n in preps[p]["nodes"].values())
        old = min(abs(delta(n, g["comp"]) / 1e6 - g["amount_m"]) for n in preps[prev]["nodes"].values())
        if old < 0.06 and cur > 0.25:
            drop.add(i)
    if drop and verbose:
        print(f"  {len(drop)} box(es) todavía con texto del mes anterior: no se usan para aprender.")
    gold = [g for i, g in enumerate(gold) if i not in drop]
    for p in bad:
        preps.pop(p, None)
        if verbose:
            print(f"  {p}: el deck todavía trae los boxes del mes anterior (en proceso): no se usa para aprender.")
    gold = [g for g in gold if g["period"] in preps]
    base = copy.deepcopy(STYLE_DEFAULT)
    s0, rows0 = learn_evaluate(base, preps, gold)
    vocab = learn_vocab([g for g in gold_all if g["box"]], preps, company)
    st1 = copy.deepcopy(base)
    st1.update(vocab)
    s1, _ = learn_evaluate(st1, preps, gold)
    best, s2 = learn_tune(st1, preps, gold)
    best, s2 = tune_per_box(best, preps, gold)
    s2, rows2 = learn_evaluate(best, preps, gold)
    # precisión honesta: cada mes se califica con un estilo aprendido SIN ese mes (leave-one-month-out)
    cv = []
    if len(preps) >= 4 and os.environ.get("VEMO_MFR_CV", "0") == "1":
        for p in sorted(preps):
            tr = {k: v for k, v in preps.items() if k != p}
            gtr = [g for g in gold if g["period"] != p]
            st_cv, _ = learn_tune(st1, tr, gtr, passes=1)
            st_cv, _ = tune_per_box(st_cv, tr, gtr)
            sc, _ = learn_evaluate(st_cv, {p: preps[p]}, [g for g in gold if g["period"] == p])
            cv.append(sc)
    s_cv = sum(cv) / len(cv) if cv else None
    set_style(best)
    res = {"company": company, "trained_on": sorted(preps), "n_boxes": len(rows2), "score_before": round(s0, 3),
           "score_cv": round(s_cv, 3) if s_cv is not None else None,
           "score_vocab": round(s1, 3), "score_after": round(s2, 3), "seconds": round(time.time() - t0, 1),
           "style": best, "when": time.strftime("%Y-%m-%d %H:%M")}
    if out_path:
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        with open(out_path, "w", encoding="utf-8") as fh:
            json.dump(res, fh, ensure_ascii=False, indent=1, default=str)
    if report_path:
        _learn_report(report_path, res, rows0, rows2)
    set_style(best)
    if verbose:
        print(f"MFR {company}: parecido con tus boxes {s0:.0%} -> {s2:.0%}"
              + (f" (en meses no vistos: {s_cv:.0%})" if s_cv is not None else "")
              + f" · {len(rows2)} boxes, meses {', '.join(sorted(preps))}.")
    return res


def score_style(style, company, fy_path, gold_path, periods):
    """Califica un estilo guardado sobre los mismos meses (para no adoptar uno peor)."""
    wb = open_wb(fy_path)
    gold = []
    for co in ("DAE", "LTO", "EVF", "VCN"):
        gold += gold_records(gold_path, co)
    gold = [x for x in gold if x["company"] == company and x["period"] in periods]
    preps = {p: prepare(wb, company, p, path=fy_path) for p in periods}
    _assign_boxes_by_header(gold, preps)
    full = copy.deepcopy(STYLE_DEFAULT)
    for k, v in style.items():
        if isinstance(v, dict) and isinstance(full.get(k), dict):
            full[k].update(v)
        else:
            full[k] = v
    return learn_evaluate(full, preps, gold)


def quality_score(company, fy_path, gold_path):
    """Qué tanto los boxes de la IA (perfil 'calidad') cubren los drivers de los boxes de referencia."""
    if not os.path.exists(gold_path):
        return None
    with open(gold_path, encoding="utf-8") as fh:
        gold = [g for g in json.load(fh) if g["company"] == company]
    if not gold:
        return None
    load_style(perfil="calidad")
    wb = open_wb(fy_path)
    scores, pers = [], sorted({g["period"] for g in gold})
    for p in pers:
        r = compose(prepare(wb, company, p, path=fy_path))
        for g in [x for x in gold if x["period"] == p]:
            b = next((x for x in r["boxes"] if x["box"] == g["box"] and not x.get("missing")), None)
            if b:
                scores.append(learn_score(g["text"], b[g["comp"]], g["box"])[0])
    return {"drivers": sum(scores) / len(scores) if scores else 0.0, "periodos": pers, "n": len(scores)}


def _learn_report(path, res, rows0, rows2):
    L = [f"# Backtest MFR {res['company']} — la IA vs. tus boxes publicados", "",
         f"Parecido promedio (drivers que tú mencionaste vs. los de la IA): **{res['score_before']:.0%} -> {res['score_after']:.0%}** "
         f"en {res['n_boxes']} boxes ({', '.join(res['trained_on'])})."
         + (f" Calificado en meses que no vio al aprender: **{res['score_cv']:.0%}**." if res.get("score_cv") is not None else ""), "",
         "| Mes | Box | Comp | Antes | Después | Tus drivers | IA | Encabezado tuyo / IA |", "|---|---|---|---|---|---|---|---|"]
    before = {(r["period"], r["box"], r["comp"]): r["learn_score"] for r in rows0}
    for r in rows2:
        L.append(f"| {r['period']} | {r['box']} | {r['comp']} | {before.get((r['period'], r['box'], r['comp']), 0):.0%} | {r['learn_score']:.0%} | "
                 f"{', '.join(r['gold_keys'])} | {', '.join(r['gen_keys'])} | {r['head_gold']} / {r['head_gen']} |")
    L += ["", "## Estilo aprendido", "```json", json.dumps({k: v for k, v in res["style"].items()}, ensure_ascii=False, indent=1, default=str), "```"]
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(L))


# ====================================================================================================
# histórico de todas las empresas (FY Financials de cada año, auditados cuando existen)
# ====================================================================================================

HIST_COMPANIES = {"DAE": r"\bdae\b(?![-_ ]?co\b)", "LTO": r"\blto\b", "EVF": r"evfleets|ev fleets", "VCN": r"\bvcn\b",
                  "CORP": r"\bcorp\b", "ZEE": r"\bzee\b"}
HIST_EXCLUDE = (r"backup|back up|trash|copia|copy|\bkari\b|quarter|financieros ly|proxy|version|_old|fv1|newv|\bbd\b|"
                r"post board|promecap|dynamics|p&l pasadas|inventarios|\(1\)|\(2\)|_2\.xlsx|pasado|_oa\b|nnnn|charly|fnmxw|"
                r"claude outputs|nuevo|jejeje|maximilian|almana")
HIST_LINES = {
    "revenue": ("revenues", "revenue", "total revenues", "total revenue"),
    "nor": ("net operating revenue", "net operating revenues"),
    "fm": ("financial margin",), "afm": ("adjusted financial margin",),
    "opex": ("opex", "total opex"), "gp": ("gross profit",),
    "dlsga": ("direct line sg&a", "direct business lines sg&a", "direct business line sg&a"),
    "sga": ("total sg&a", "sg&a"), "ebitda": ("ebitda",),
    "ni": ("net income", "net income (loss)", "net loss"),
    "fleet": ("total fleet",), "drivers": ("eop active drivers",), "sh": ("supply hours",),
    "portfolio": ("gross portfolio (accounting)", "gross portfolio (risk)", "gross portfolio"),
    "units": ("active units eop", "active units"), "origination": ("origination",),
}
HIST_STOCK = {"fleet", "drivers", "portfolio", "units"}
HIST_NAMES = {"revenue": "Revenues", "nor": "Net Operating Revenue", "fm": "Financial Margin", "afm": "Adj. Financial Margin",
              "opex": "OPEX", "gp": "Gross Profit", "dlsga": "DL SG&A", "sga": "SG&A", "ebitda": "EBITDA", "ni": "Net Income",
              "fleet": "Flota total", "drivers": "Drivers activos", "sh": "Supply hours", "portfolio": "Cartera bruta",
              "units": "Unidades activas", "origination": "Originación"}


def hist_files(root):
    """{(empresa, año): ruta} del FY Financials canónico: auditado > versión final en la carpeta de su año."""
    best = {}
    base = os.path.join(root, "1. Fundraise")
    if os.path.isdir(base):  # los FY Financials viven aquí: recorrer solo esta carpeta es mucho más rápido
        items = []
        for dp, _dn, fn_ in os.walk(base):
            for f in fn_:
                if f.lower().endswith(".xlsx") and not f.startswith("~$"):
                    try:
                        items.append((os.path.join(dp, f), os.path.getmtime(os.path.join(dp, f)), 0))
                    except OSError:
                        pass
    else:
        items = scan_tree(root)
    for path, mt, _sz in items:
        fn = os.path.basename(path)
        pl = _plain(path.replace("\\", "/"))
        if not fn.lower().endswith(".xlsx") or "fy financials" not in pl or re.search(HIST_EXCLUDE, pl):
            continue
        m = re.search(r"vemo (20\d\d)", _plain(fn))
        if not m:
            continue
        y = int(m.group(1))
        for co, rx in HIST_COMPANIES.items():
            if re.search(rx, _plain(fn)):
                score = (2 * bool(re.search(r"audit", pl)) + (f"/{y}/" in pl.replace("\\", "/")), mt)
                if (co, y) not in best or score > best[(co, y)][0]:
                    best[(co, y)] = (score, path)
                break
    return {k: v[1] for k, v in best.items()}


def _hist_sheet_rank(title):
    t = _plain(title)
    if "board view" in t and not re.search(r"budget|\(b\)", t):
        return 0 if re.search(r"actuals|auditad", t) else (1 if "fp&a" in t else (2 if "(2)" not in t else 4))
    if t == "board tabs outputs":
        return 3
    if t.startswith("managerial consolidated") and not re.search(r"acc|comp", t):
        return 5
    return None


def _hist_read(ws, year, budget=False):
    """Serie mensual de los renglones clave (etiqueta exacta, primera aparición) para `year`."""
    import datetime as _d
    rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row or 900, 900), values_only=True))
    best = None
    for r, row in enumerate(rows[:45]):
        cols = {c: f"{v.year:04d}-{v.month:02d}" for c, v in enumerate(row) if isinstance(v, (_d.datetime, _d.date))}
        if len(cols) >= 6 and (best is None or len(cols) > len(best[1])):
            best = (r, cols)
    if not best:
        return None
    hr, cols = best
    mc = {}
    for c, ym in sorted(cols.items()):
        if ym.startswith(str(year)) and ym not in mc.values():
            mc[c] = ym
    if len(mc) < 1:
        return None
    names = {n: k for k, ns in HIST_LINES.items() for n in ns}
    occ = {}
    for i, row in enumerate(rows[hr + 1:]):
        lab = next((_plain(v) for v in row[:6] if isinstance(v, str) and v.strip()), None)
        k = names.get(lab)
        if not k:
            continue
        vals = {ym: float(row[c]) for c, ym in mc.items() if c < len(row) and isinstance(row[c], (int, float))
                and row[c] != 0 and (budget or ym <= _last_closed())}  # meses aún no cerrados no cuentan
        if vals:
            occ.setdefault(k, []).append((i, vals))
    out = {}
    # renglón superior del P&L: el de mayor monto (el mismo nombre aparece en KPIs y desgloses)
    top_k = "nor" if "nor" in occ else ("revenue" if "revenue" in occ else None)
    r0 = 0
    if top_k:
        r0, out[top_k] = max(occ[top_k], key=lambda t: abs(sum(t[1].values())))
        if top_k == "nor" and "revenue" in occ:
            out["revenue"] = max(occ["revenue"], key=lambda t: abs(sum(t[1].values())))[1]
    for k, cands in occ.items():
        if k in out:
            continue
        if k in HIST_STOCK or k in ("sh", "origination"):
            out[k] = cands[0][1]  # KPIs: el bloque de KPIs (arriba)
            continue
        after = [t for t in cands if t[0] > r0]
        if k in ("ebitda", "ni", "sga", "dlsga"):  # el primero debajo del renglón superior del P&L
            out[k] = (after or cands)[0][1]
        else:
            out[k] = max(after or cands, key=lambda t: abs(sum(t[1].values())))[1]
    if "opex" in occ and "revenue" in out and "gp" in out:  # si hay varios OPEX, el que cuadra Revenues + OPEX = GP
        r, g = sum(out["revenue"].values()), sum(out["gp"].values())
        out["opex"] = min((t[1] for t in occ["opex"]), key=lambda d: abs(r + sum(d.values()) - g))
    return out


def _hist_ok(d):
    """0 = no sirve (sin ingresos o sin utilidad neta); 1 = legible; 2 = además cuadra Revenues + OPEX = Gross Profit (±1%)."""
    if not d:
        return 0
    top = d.get("revenue") or d.get("nor")
    if not top or "ni" not in d or sum(1 for v in top.values() if v) < 1:
        return 0
    if all(k in d for k in ("revenue", "opex", "gp")):
        r, o, g = (sum(d[k].values()) for k in ("revenue", "opex", "gp"))
        return 2 if abs(r + o - g) <= max(0.01 * abs(r), 0.5e6) else 1
    return 1


def hist_extract(path, year):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    act = bud = None
    ranked = sorted([(r, ws.title) for ws in wb.worksheets if (r := _hist_sheet_rank(ws.title)) is not None])
    best = None
    for r, t in ranked:  # mejor hoja: la que cuadra; a igualdad, la de más prioridad (Board View) y más meses
        d = _hist_read(wb[t], year)
        ok = _hist_ok(d)
        if not ok:
            continue
        n = len(d.get("revenue") or d.get("nor") or {})
        key = (r <= 3, n, ok, -r)  # el Board View (lo que va al MFR) manda; Managerial solo si no hay
        if best is None or key > best[0]:
            best = (key, {"sheet": t, "lines": d, "cuadra": ok == 2})
    act = best[1] if best else None
    for ws in wb.worksheets:
        if re.search(r"board view.*(budget|\(b\))", _plain(ws.title)):
            d = _hist_read(ws, year, budget=True)
            if _hist_ok(d):
                bud = {"sheet": ws.title, "lines": d}
                break
    wb.close()
    return act, bud


def build_history(root, verbose=True):
    """datos/historico.json: P&L mensual y KPIs de cada empresa y año. Solo relee los archivos que cambiaron."""
    hp = os.path.join(KNOWLEDGE_DIR, "historico.json")
    try:
        with open(hp, encoding="utf-8") as fh:
            H = json.load(fh)
    except Exception:
        H = {}
    files = hist_files(root)
    changed = False
    for (co, y), path in sorted(files.items()):
        try:
            mt = int(os.path.getmtime(path))
        except OSError:
            continue
        cur = H.get(co, {}).get(str(y))
        rel = os.path.relpath(path, root).replace("\\", "/")
        if cur and cur.get("file") == rel and cur.get("mtime") == mt:
            continue
        try:
            act, bud = hist_extract(path, y)
        except Exception as e:
            if verbose:
                print(f"  !! {co} {y}: {e}")
            continue
        H.setdefault(co, {})[str(y)] = {"file": rel, "mtime": mt, "audit": bool(re.search(r"audit", _plain(rel))),
                                         "actual": act, "budget": bud}
        changed = True
        with open(hp, "w", encoding="utf-8") as fh:  # se guarda archivo por archivo (si se interrumpe, retoma)
            json.dump(H, fh, ensure_ascii=False)
        if verbose:
            top = (act or {}).get("lines", {})
            t = top.get("revenue") or top.get("nor") or {}
            print(f"  {co} {y}: {'OK' if act else 'sin P&L legible'} · {len(t)} meses · {os.path.basename(path)}"
                  + (f" [{act['sheet']}]" if act else ""), flush=True)
    if changed or not os.path.exists(hp):
        with open(hp, "w", encoding="utf-8") as fh:
            json.dump(H, fh, ensure_ascii=False)
    return H


def load_history():
    try:
        with open(os.path.join(KNOWLEDGE_DIR, "historico.json"), encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def _hsum(series, k, months=None):
    if not series:
        return None
    vals = {ym: v for ym, v in series.items() if months is None or int(ym[5:7]) in months}
    if not vals:
        return None
    if k in HIST_STOCK:
        return vals[max(vals)]
    return sum(vals.values())


def _hfmt(v, k):
    if v is None:
        return "n.d."
    if k in ("fleet", "drivers", "units", "origination"):
        return f"{v:,.0f}"
    if k == "sh":
        return f"{v / 1000:,.0f}k"
    return _m(v)


def history_md(co, per, cur_rows=None):
    """Tabla anual por año + comparación contra el mismo periodo del año anterior."""
    H = load_history().get(co) or {}
    if not H:
        return []
    yrs = sorted(int(y) for y in H if (H[y].get("actual") or {}).get("lines"))
    if not yrs:
        return []
    has_nor = any("nor" in H[str(y)]["actual"]["lines"] for y in yrs)
    keys = [k for k in ("revenue", "nor", "gp", "afm", "ebitda", "ni") if any(k in H[str(y)]["actual"]["lines"] for y in yrs)
            and not (k == "revenue" and has_nor)]
    L = ["", f"## Histórico {co} (FY Financials, MXN m)", "",
         "| Año | " + " | ".join(HIST_NAMES[k] for k in keys) + " | Meses | Fuente | Cuadre |",
         "|---|" + "---|" * (len(keys) + 3)]
    for y in yrs:
        a = H[str(y)]["actual"]
        top = a["lines"].get("nor") or a["lines"].get("revenue") or {}
        L.append(f"| {y} | " + " | ".join(_hfmt(_hsum(a["lines"].get(k), k), k) for k in keys)
                 + f" | {len(top)} | {'auditado' if H[str(y)]['audit'] else 'FV'} · {a['sheet']} | "
                 + ("Rev+OPEX=GP ✓" if a.get("cuadra") else "revisar") + " |")
    y, m = int(per[:4]), int(per[5:7])
    ly = H.get(str(y - 1), {}).get("actual")
    if ly:
        rng = set(range(1, m + 1))
        now = (H.get(str(y)) or {}).get("actual") or {}
        L += ["", f"**{MES_ES[0]}–{MES_ES[m - 1]} {y} vs. {y - 1}:** " + "; ".join(
            f"{HIST_NAMES[k]} {_m(_hsum(now['lines'].get(k), k, rng))} vs. {_m(_hsum(ly['lines'].get(k), k, rng))}"
            for k in keys if now.get("lines", {}).get(k) and ly["lines"].get(k)) + "."]
        same = [(k, now["lines"][k].get(per), ly["lines"][k].get(f"{y - 1}-{m:02d}")) for k in keys
                if now.get("lines", {}).get(k) and ly["lines"].get(k)]
        if same:
            L.append(f"**{MES_ES[m - 1]} {y} vs. {MES_ES[m - 1]} {y - 1}:** " + "; ".join(
                f"{HIST_NAMES[k]} {_m(a)} vs. {_m(b)}" for k, a, b in same if a is not None and b is not None) + ".")
    kk = [k for k in ("fleet", "drivers", "sh", "portfolio", "units", "origination") if any(k in H[str(y)]["actual"]["lines"] for y in yrs)]
    if kk:
        L += ["", "Operación al cierre de cada año (o al último mes cerrado): " + "; ".join(
            f"{y}: " + ", ".join(f"{HIST_NAMES[k]} {_hfmt(_hsum(H[str(y)]['actual']['lines'].get(k), k), k)}"
                               for k in kk if H[str(y)]["actual"]["lines"].get(k))
            for y in yrs if any(H[str(y)]["actual"]["lines"].get(k) for k in kk)) + "."]
    return L


# ====================================================================================================
# Dynamics (DB.xlsx: pólizas del mayor) -> qué movimientos hay detrás de cada renglón del Board View
# ====================================================================================================

# ----------------------------------------------------------------------------------------------------
# Dynamics 365 Business Central por API (OData V4). El secreto NUNCA va a OneDrive ni al repo:
# se guarda en el Administrador de credenciales de Windows (o en ~/.vemo_ia fuera de OneDrive).
# ----------------------------------------------------------------------------------------------------

BC_DEFAULT = {"tenant": "7f4043a9-2a49-4a83-a6a7-56fcf6fa8d04", "entorno": "Production",
              "empresa": "MOBILITY DAE PRODUCCION", "endpoint": "Movs_contabilidad_Excel"}
BC_SCOPE = "https://api.businesscentral.dynamics.com/.default"
BC_FIELDS = {  # nombre lógico -> nombres posibles en la API / en el Excel
    "date": ("Posting_Date", "PostingDate", "postingDate"),
    "acct": ("G_L_Account_No", "GLAccountNo", "accountNumber", "G_L_Account_No_"),
    "name": ("G_L_Account_Name", "GLAccountName", "accountName"),
    "desc": ("Description", "description"),
    "amt": ("Amount", "amount"),
    "doc": ("Document_No", "DocumentNo", "documentNumber"),
}


def _cred_target(name):
    return f"vemo_ia/{name}"


def secret_set(name, value):
    if os.name == "nt":
        import ctypes
        import ctypes.wintypes as wt

        class CRED(ctypes.Structure):
            _fields_ = [("Flags", wt.DWORD), ("Type", wt.DWORD), ("TargetName", wt.LPWSTR), ("Comment", wt.LPWSTR),
                        ("LastWritten", wt.FILETIME), ("CredentialBlobSize", wt.DWORD),
                        ("CredentialBlob", ctypes.POINTER(ctypes.c_char)), ("Persist", wt.DWORD),
                        ("AttributeCount", wt.DWORD), ("Attributes", ctypes.c_void_p), ("TargetAlias", wt.LPWSTR),
                        ("UserName", wt.LPWSTR)]
        blob = value.encode("utf-8")
        c = CRED()
        c.Type, c.Persist = 1, 2  # GENERIC, LOCAL_MACHINE (solo este usuario de Windows)
        c.TargetName, c.UserName = _cred_target(name), "vemo_ia"
        c.CredentialBlobSize = len(blob)
        c.CredentialBlob = ctypes.cast(ctypes.create_string_buffer(blob, len(blob)), ctypes.POINTER(ctypes.c_char))
        if not ctypes.windll.advapi32.CredWriteW(ctypes.byref(c), 0):
            raise OSError("no pude guardar la credencial en Windows")
        return
    d = os.path.join(os.path.expanduser("~"), ".vemo_ia")
    os.makedirs(d, exist_ok=True)
    p = os.path.join(d, name)
    with open(p, "w", encoding="utf-8") as fh:
        fh.write(value)
    os.chmod(p, 0o600)


def secret_get(name):
    env = os.environ.get("VEMO_BC_" + name.upper())
    if env:
        return env
    if os.name == "nt":
        import ctypes
        import ctypes.wintypes as wt

        class CRED(ctypes.Structure):
            _fields_ = [("Flags", wt.DWORD), ("Type", wt.DWORD), ("TargetName", wt.LPWSTR), ("Comment", wt.LPWSTR),
                        ("LastWritten", wt.FILETIME), ("CredentialBlobSize", wt.DWORD),
                        ("CredentialBlob", ctypes.POINTER(ctypes.c_char)), ("Persist", wt.DWORD),
                        ("AttributeCount", wt.DWORD), ("Attributes", ctypes.c_void_p), ("TargetAlias", wt.LPWSTR),
                        ("UserName", wt.LPWSTR)]
        p = ctypes.POINTER(CRED)()
        if not ctypes.windll.advapi32.CredReadW(_cred_target(name), 1, 0, ctypes.byref(p)):
            return None
        try:
            return ctypes.string_at(p.contents.CredentialBlob, p.contents.CredentialBlobSize).decode("utf-8")
        finally:
            ctypes.windll.advapi32.CredFree(p)
    try:
        with open(os.path.join(os.path.expanduser("~"), ".vemo_ia", name), encoding="utf-8") as fh:
            return fh.read().strip()
    except OSError:
        return None


def dyn_api_cfg(company="DAE"):
    try:
        with open(os.path.join(KNOWLEDGE_DIR, "dynamics_api.json"), encoding="utf-8") as fh:
            cfg = json.load(fh).get(company)
        return cfg if cfg and cfg.get("activo") else None
    except Exception:
        return None


def _http(url, data=None, headers=None, timeout=90, retries=4):
    import urllib.error
    import urllib.request
    body = urllib.parse.urlencode(data).encode() if isinstance(data, dict) else data
    for i in range(retries):
        req = urllib.request.Request(url, data=body, headers=headers or {})
        try:
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            txt = e.read().decode("utf-8", "ignore")[:400]
            if e.code in (429, 500, 502, 503, 504) and i < retries - 1:
                time.sleep(int(e.headers.get("Retry-After") or 2 ** (i + 1)))
                continue
            try:
                j = json.loads(txt)
                err = j.get("error")
                msg = j.get("error_description") or (err.get("message") if isinstance(err, dict) else err) or txt
            except Exception:
                msg = txt
            raise RuntimeError(f"HTTP {e.code}: {msg}") from None
        except urllib.error.URLError as e:
            if i < retries - 1:
                time.sleep(2 ** (i + 1))
                continue
            raise RuntimeError(f"sin conexión con {urllib.parse.urlparse(url).netloc}: {e.reason}") from None


class BCClient:
    """Cliente mínimo de Business Central (OData V4): token de Entra ID + paginación @odata.nextLink."""

    def __init__(self, cfg):
        self.cfg = {**BC_DEFAULT, **cfg}
        self._tok, self._exp = None, 0
        self.login = self.cfg.get("login_url") or f"https://login.microsoftonline.com/{self.cfg['tenant']}/oauth2/v2.0"
        base = self.cfg.get("api_base") or "https://api.businesscentral.dynamics.com/v2.0"
        comp = urllib.parse.quote(self.cfg["empresa"], safe="")
        self.url = f"{base}/{self.cfg['tenant']}/{self.cfg['entorno']}/ODataV4/Company('{comp}')/{self.cfg['endpoint']}"

    def token(self):
        if self._tok and time.time() < self._exp - 120:
            return self._tok
        c = self.cfg
        if c.get("auth", "client_credentials") == "client_credentials":
            sec = secret_get("client_secret")
            if not sec:
                raise RuntimeError("falta el secreto de la app (corre: vemo_ia.py dynamics --configurar)")
            j = _http(f"{self.login}/token", {"grant_type": "client_credentials", "client_id": c["client_id"],
                                               "client_secret": sec, "scope": BC_SCOPE})
        else:
            rt = secret_get("refresh_token")
            if not rt:
                raise RuntimeError("no hay sesión iniciada (corre: vemo_ia.py dynamics --configurar)")
            j = _http(f"{self.login}/token", {"grant_type": "refresh_token", "client_id": c["client_id"],
                                               "refresh_token": rt, "scope": BC_SCOPE + " offline_access"})
            if j.get("refresh_token"):
                secret_set("refresh_token", j["refresh_token"])
        self._tok, self._exp = j["access_token"], time.time() + int(j.get("expires_in", 3600))
        return self._tok

    def device_login(self):
        c = self.cfg
        j = _http(f"{self.login}/devicecode", {"client_id": c["client_id"], "scope": BC_SCOPE + " offline_access"})
        print("\n" + j.get("message", f"Entra a {j['verification_uri']} y escribe el código {j['user_code']}"))
        t0, wait = time.time(), int(j.get("interval", 5))
        while time.time() - t0 < int(j.get("expires_in", 900)):
            time.sleep(wait)
            try:
                tok = _http(f"{self.login}/token", {"grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                                                     "client_id": c["client_id"], "device_code": j["device_code"]}, retries=1)
            except RuntimeError as e:
                if "authorization_pending" in str(e) or "AADSTS70016" in str(e) or "pending" in str(e).lower():
                    continue
                if "slow_down" in str(e):
                    wait += 5
                    continue
                raise
            secret_set("refresh_token", tok["refresh_token"])
            return True
        raise RuntimeError("se venció el código sin iniciar sesión")

    def rows(self, flt=None, select=None, top=None):
        q = {}
        if flt:
            q["$filter"] = flt
        if select:
            q["$select"] = select
        if top:
            q["$top"] = str(top)
        url = self.url + ("?" + urllib.parse.urlencode(q, quote_via=urllib.parse.quote) if q else "")
        while url:
            j = _http(url, headers={"Authorization": f"Bearer {self.token()}", "Accept": "application/json",
                                    "Prefer": "odata.maxpagesize=20000"})
            yield from j.get("value", [])
            url = j.get("@odata.nextLink")


def _f(r, k):
    for n in BC_FIELDS[k]:
        if n in r:
            return r[n]
    return None


def _dyn_agg(rows, mp, agg=None):
    """Agrega pólizas (dicts) por renglón del Board y mes. Signo del Board: ingreso +, gasto −."""
    A = agg or {"tot": {}, "acc": {}, "top": {}, "cnt": {}, "unm": Counter(), "ultimo": ""}
    for r in rows:
        a = str(_f(r, "acct") or "").strip()
        if not a or a[:1] not in "4567":
            continue
        d = _f(r, "date")
        ds = d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d)[:10]
        ym = ds[:7]
        try:
            amt = -float(_f(r, "amt") or 0)
        except (TypeError, ValueError):
            continue
        line = mp.get(a)
        if not line:
            A["unm"][a] += 1
            continue
        key = f"{'|'.join(line)}#{ym}"
        A["tot"][key] = A["tot"].get(key, 0.0) + amt
        A["cnt"][key] = A["cnt"].get(key, 0) + 1
        an = str(_f(r, "name") or a).strip()
        A["acc"].setdefault(key, {})
        A["acc"][key][an] = A["acc"][key].get(an, 0.0) + amt
        desc = re.sub(r"\s+", " ", str(_f(r, "desc") or an)).strip()[:90]
        x = A["top"].setdefault(key, {}).setdefault(desc, [0.0, an, ds, str(_f(r, "doc") or ""), 0])
        x[0] += amt
        x[4] += 1
        A["ultimo"] = max(A["ultimo"], ds)
    return A


def _dyn_finish(A):
    top = {k: [[round(v[0], 2), dsc, v[1], v[2], v[3], v[4]] for dsc, v in
               sorted(t.items(), key=lambda it: -abs(it[1][0])) if abs(v[0]) >= 1000][:8] for k, t in A["top"].items()}
    return {"totales": A["tot"], "cuentas": A["acc"], "top": top, "n": A["cnt"], "sin_mapeo": len(A["unm"]),
            "ultimo_mov": A["ultimo"], "meses": sorted({k.split("#")[1] for k in A["tot"]})}


def _next_month(ym):
    y, m = int(ym[:4]), int(ym[5:7]) + 1
    return f"{y + (m > 12):04d}-{(m - 1) % 12 + 1:02d}"


def dyn_api_fetch(cfg, mp, months, verbose=True):
    """Trae de la API las pólizas de `months` (YYYY-MM) y las agrega. Un mes por llamada (filtro por fecha)."""
    cli = BCClient(cfg)
    A = None
    sel = ",".join(BC_FIELDS[k][0] for k in ("date", "acct", "name", "desc", "amt", "doc"))
    for ym in months:
        flt = f"Posting_Date ge {ym}-01 and Posting_Date lt {_next_month(ym)}-01"
        try:
            rows = list(cli.rows(flt, sel))
        except RuntimeError as e:
            if "HTTP 400" in str(e):  # el servicio no acepta $select con esos nombres: sin $select
                rows = list(cli.rows(flt))
            else:
                raise
        A = _dyn_agg(rows, mp, A)
        if verbose:
            print(f"    {ym}: {len(rows):,} pólizas", flush=True)
    return A or _dyn_agg([], mp)


def _dyn_merge(old, new, months):
    """Reemplaza en `old` los meses `months` con lo de `new`."""
    out = {k: dict(old.get(k) or {}) for k in ("totales", "cuentas", "top", "n")}
    for k in out:
        for key in [x for x in out[k] if x.split("#")[1] in months]:
            out[k].pop(key)
        out[k].update(new.get(k) or {})
    out["meses"] = sorted({k.split("#")[1] for k in out["totales"]})
    out["ultimo_mov"] = max(old.get("ultimo_mov") or "", new.get("ultimo_mov") or "")
    out["sin_mapeo"] = new.get("sin_mapeo", 0)
    return out


def dyn_api_configurar(company="DAE"):
    """Asistente: guarda la conexión en datos/dynamics_api.json (sin secretos) y el secreto en Windows."""
    p = os.path.join(KNOWLEDGE_DIR, "dynamics_api.json")
    try:
        with open(p, encoding="utf-8") as fh:
            allc = json.load(fh)
    except Exception:
        allc = {}
    cfg = {**BC_DEFAULT, **(allc.get(company) or {})}
    print("\nConexión a Dynamics 365 Business Central (API). Enter = dejar lo que está entre [].")
    link = input("Pega el link de la API (o Enter) : ").strip()
    m = re.search(r"/v2\.0/([0-9a-f-]{36})/([^/]+)/ODataV4/Company\('([^']+)'\)/([^/?]+)", urllib.parse.unquote(link))
    if m:
        cfg.update(tenant=m.group(1), entorno=m.group(2), empresa=m.group(3), endpoint=m.group(4))
    print(f"  Tenant {cfg['tenant']} · {cfg['entorno']} · {cfg['empresa']} · {cfg['endpoint']}")
    print("  Tipo de acceso:\n   1) App de Entra ID con secreto (TI te da Client ID y Secret)  [recomendado]\n"
          "   2) Iniciar sesión con tu cuenta de Microsoft (TI te da el Client ID de una app pública)")
    op = input(f"  Opción [{1 if cfg.get('auth', 'client_credentials') == 'client_credentials' else 2}]: ").strip()
    cfg["auth"] = "device_code" if op == "2" or (not op and cfg.get("auth") == "device_code") else "client_credentials"
    cid = input(f"  Client ID [{cfg.get('client_id', '')}]: ").strip()
    if cid:
        cfg["client_id"] = cid
    if not cfg.get("client_id"):
        print("  Sin Client ID no puedo conectarme. Pídeselo a TI (ver README) y vuelve a correr esto.")
        return False
    if cfg["auth"] == "client_credentials":
        import getpass
        sec = getpass.getpass("  Client Secret (no se muestra; Enter = conservar el guardado): ").strip()
        if sec:
            secret_set("client_secret", sec)
    sa = input(f"  ¿Usar SOLO la API (no leer DB.xlsx)? s/n [{'s' if cfg.get('solo_api', True) else 'n'}]: ").strip().lower()
    cfg["solo_api"] = (sa != "n") if sa else cfg.get("solo_api", True)
    cfg["activo"] = True
    allc[company] = {k: v for k, v in cfg.items() if k not in ("login_url", "api_base") or v}
    os.makedirs(KNOWLEDGE_DIR, exist_ok=True)
    with open(p, "w", encoding="utf-8") as fh:
        json.dump(allc, fh, ensure_ascii=False, indent=1)
    if cfg["auth"] == "device_code":
        BCClient(cfg).device_login()
    return dyn_api_probar(company)


def dyn_api_probar(company="DAE"):
    cfg = dyn_api_cfg(company)
    if not cfg:
        print("  Dynamics API no configurada (vemo_ia.py dynamics --configurar).")
        return False
    try:
        cli = BCClient(cfg)
        r = next(cli.rows(top=1), None)
    except Exception as e:
        print(f"  ✗ No pude conectarme: {e}")
        return False
    if r is None:
        print("  ✓ Conectado, pero el servicio no regresó pólizas.")
        return True
    falt = [k for k in BC_FIELDS if _f(r, k) is None and k not in ("doc",)]
    print(f"  ✓ Conectado a {cfg.get('empresa', BC_DEFAULT['empresa'])}: {len(r)} campos por póliza"
          + (f"; OJO, no encuentro: {', '.join(falt)}" if falt else "; campos completos."))
    return not falt


def dyn_db_path(fy_path):
    """El DB.xlsx que extraes de Dynamics, en la carpeta del FY Financials (p. ej. '13. DB Dynamics/DB.xlsx')."""
    base = os.path.dirname(fy_path)
    cands = []
    for dp, _dn, fn in os.walk(base):
        if "dynamics" not in _plain(dp) or re.search(r"backup|respaldo", _plain(dp)):
            continue
        for f in fn:
            if f.lower().endswith(".xlsx") and not f.startswith("~$") and re.match(r"(?i)db\b", f):
                cands.append(os.path.join(dp, f))
    return max(cands, key=os.path.getmtime) if cands else None


def dyn_mapping(wb):
    """Cuenta contable -> (renglón del Board, área, sub) desde la hoja 'TB DATA BASE' del FY Financials."""
    if "TB DATA BASE" not in wb.sheetnames:
        return {}
    ws = wb["TB DATA BASE"]
    mp, cols = {}, None
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row or 2000, 2000), values_only=True):
        if cols is None:
            labs = [_plain(x) if isinstance(x, str) else "" for x in row]
            if "account number" in labs:
                ix = {v: i for i, v in enumerate(labs)}
                cols = (ix["account number"], ix.get("fs budget (board)"), ix.get("fs budget (board) l-2"),
                        ix.get("fs budget (board) l-3"), ix.get("fs budget (board) l-4"))
            continue
        a = row[cols[0]] if cols[0] < len(row) else None
        if not a or cols[2] is None:
            continue
        g = lambda i: (str(row[i]).strip() if i is not None and i < len(row) and row[i] not in (None, "n.a.") else "")
        if g(cols[2]):
            new = (g(cols[1]), g(cols[2]), g(cols[3]), g(cols[4]))
            key = str(a).strip()
            spec = lambda t: sum(1 for x in t[2:] if x and _plain(x) != _plain(t[1]))  # sub-renglón más específico
            if key not in mp or spec(new) > spec(mp[key]):  # la cuenta viene repetida: me quedo con el mapeo más fino
                mp[key] = new
    return mp


def dyn_build(root, company="DAE", verbose=True, force=False):
    """datos/dynamics_<CO>.json: por renglón del Board y mes, total, desglose por cuenta y movimientos más grandes.
    Fuente: la API de Business Central si está configurada (incremental: refresca los últimos meses);
    si no, el DB.xlsx que exportas. Con 'solo_api' nunca usa el Excel."""
    import openpyxl
    fy = fy_file(root, company, _dt.date.today().year)
    if not fy:
        return None
    out_p = os.path.join(KNOWLEDGE_DIR, f"dynamics_{company}.json")
    try:
        with open(out_p, encoding="utf-8") as fh:
            old = json.load(fh)
    except Exception:
        old = None
    cfg = dyn_api_cfg(company)
    fy_mt = int(os.path.getmtime(fy))
    if cfg:
        if old and old.get("fuente") == "api" and old.get("fy_mtime") == fy_mt and not force:
            age = time.time() - old.get("ts", 0)
            if age < 30 * 60:  # menos de 30 min: no vuelvo a pedir
                return old
        mp = dyn_mapping(openpyxl.load_workbook(fy, read_only=True, data_only=True))
        today = _dt.date.today()
        cur = f"{today.year:04d}-{today.month:02d}"
        full = not (old and old.get("fuente") == "api" and old.get("fy_mtime") == fy_mt)
        start = _prev(f"{today.year:04d}-01") if full else _prev(_prev(cur))  # completo: desde diciembre del año pasado
        months, m = [], start
        while m <= cur:
            months.append(m)
            m = _next_month(m)
        try:
            if verbose:
                print(f"  Dynamics API: {'carga completa' if full else 'actualizo'} {months[0]}..{months[-1]} ...", flush=True)
            new = _dyn_finish(dyn_api_fetch(cfg, mp, months, verbose=verbose))
            data = new if full else _dyn_merge(old, new, months)
            data.update(fuente="api", db=f"API {cfg.get('empresa', BC_DEFAULT['empresa'])}", fy_mtime=fy_mt, ts=time.time(),
                        actualizado=_dt.datetime.now().strftime("%Y-%m-%d %H:%M"))
            with open(out_p, "w", encoding="utf-8") as fh:
                json.dump(data, fh, ensure_ascii=False)
            if verbose:
                print(f"  Dynamics API: {sum(data['n'].values()):,} pólizas de P&L, {data['meses'][0]}..{data['meses'][-1]}, "
                      f"último movimiento {data['ultimo_mov']}", flush=True)
            return data
        except Exception as e:
            msg = f"Dynamics API falló ({e})"
            if cfg.get("solo_api", True):
                if verbose:
                    print(f"  {msg}; uso lo último que bajé" + (f" ({old.get('actualizado')})" if old else " (no hay nada guardado)"))
                if old:
                    old["aviso"] = msg
                return old
            if verbose:
                print(f"  {msg}; uso DB.xlsx")
    db = dyn_db_path(fy)
    if not db:
        return old
    sig = ["v3", os.path.basename(db), int(os.path.getmtime(db)), fy_mt]
    if old and old.get("sig") == sig and not force:
        return old
    if verbose:
        print(f"  Leyendo Dynamics ({os.path.basename(db)}) ...", flush=True)
    mp = dyn_mapping(openpyxl.load_workbook(fy, read_only=True, data_only=True))
    wb = openpyxl.load_workbook(db, read_only=True, data_only=True)
    it = wb.worksheets[0].iter_rows(values_only=True)
    hdr = [str(h) if h is not None else "" for h in next(it)]
    if not all(any(n in hdr for n in BC_FIELDS[k]) for k in ("date", "acct", "amt")):
        return old
    data = _dyn_finish(_dyn_agg((dict(zip(hdr, r)) for r in it), mp))
    data.update(sig=sig, fuente="excel", db=os.path.relpath(db, root).replace("\\", "/"), fy_mtime=fy_mt, ts=time.time(),
                actualizado=_dt.datetime.fromtimestamp(os.path.getmtime(db)).strftime("%Y-%m-%d %H:%M"))
    with open(out_p, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False)
    if verbose:
        print(f"  Dynamics (Excel): {sum(data['n'].values()):,} pólizas de P&L, {data['meses'][0]}..{data['meses'][-1]}"
              + (f", {data['sin_mapeo']} cuentas sin mapeo en TB DATA BASE" if data["sin_mapeo"] else ""), flush=True)
    return data


def load_dyn(company="DAE"):
    try:
        with open(os.path.join(KNOWLEDGE_DIR, f"dynamics_{company}.json"), encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return None


def _dyn_keys(D, label, ancestors=()):
    """Claves 'L1|L2|L3|L4' del mapeo que corresponden a un renglón del Board. `ancestors`: etiquetas de sus padres
    (para separar, p. ej., Insurance de OPEX vs. Insurance de Direct Line SG&A)."""
    lab = _plain(label)
    anc = [_plain(a) for a in ancestors if a]
    ks = [k.split("|") for k in {k.split("#")[0] for k in D["totales"]}]
    in_anc = lambda l1: not anc or not l1 or _plain(l1) in anc or (_plain(l1) == "revenue" and any("revenue" in a for a in anc))
    def pick(c):  # el filtro por ancestro solo hace falta si el mismo renglón existe en dos bloques (OPEX vs. DL SG&A)
        if len({k[0] for k in c}) > 1:
            c = [k for k in c if in_anc(k[0])] or c
        return c
    own = pick([k for k in ks if _plain(k[1]) == lab])
    if not own and anc:
        own = pick([k for k in ks if _plain(k[1]) == anc[0] and lab in (_plain(k[2]), _plain(k[3]))])
    return ["|".join(k) for k in own]


def _ancestors(prep, node):
    parent = {}

    def walk(ns, p=None):
        for x in ns:
            parent[id(x)] = p
            walk(x["children"], x)
    walk(prep["roots"])
    out, p = [], parent.get(id(node))
    while p is not None:
        out.append(p["label"])
        p = parent.get(id(p))
    return out


def _amt(v):
    a = abs(v)
    s = "-" if v < 0 else "+"
    return f"{s}${a / 1e6:,.1f}m" if a >= 0.1e6 else f"{s}${a / 1e3:,.0f}k"


def dyn_line(D, keys, ym):
    if not D or not keys:
        return None
    tot = sum(D["totales"].get(f"{k}#{ym}", 0.0) for k in keys)
    n = sum(D["n"].get(f"{k}#{ym}", 0) for k in keys)
    acc = Counter()
    for k in keys:
        for a, v in (D["cuentas"].get(f"{k}#{ym}") or {}).items():
            acc[a] += v
    top = sorted([t for k in keys for t in D["top"].get(f"{k}#{ym}", [])], key=lambda t: -abs(t[0]))[:5]
    return {"total": tot, "n": n, "cuentas": acc, "top": top}


def dyn_md(prep, ctx, company="DAE", k=10):
    """Sección 'Lo que dice Dynamics' para los renglones que más se movieron."""
    D = load_dyn(company)
    per = prep["period"]
    if not D or per not in D["meses"]:
        return []
    prev = _prev(per)
    src = "API Business Central" if D.get("fuente") == "api" else D["db"].split("/")[-1]
    L = ["", f"## Lo que dice Dynamics ({src}, actualizado {D.get('actualizado', 'n.d.')}; último movimiento {D.get('ultimo_mov', 'n.d.')})", ""]
    if D.get("aviso"):
        L += [f"_{D['aviso']}: estos datos pueden estar atrasados._", ""]
    shown = 0
    for x in ctx["lineas"]:
        n = x["node"]
        keys = _dyn_keys(D, n["label"], _ancestors(prep, n))
        cur = dyn_line(D, keys, per)
        if not cur or not cur["n"]:
            continue
        old = dyn_line(D, keys, prev) if prev in D["meses"] else None
        board = n["m"] or 0.0
        dif = board - cur["total"]
        head = (f"- **{x['label']}**: Dynamics {_m(cur['total'])} en {cur['n']} pólizas · Board {_m(board)}"
                + (f" (cuadra)" if abs(dif) <= 0.05e6 else f" (diferencia {_amt(dif)}: provisión, reclasificación o asiento fuera de este mayor; revisar)"))
        L.append(head)
        if old and old["n"]:
            dif = Counter({a: cur["cuentas"].get(a, 0) - old["cuentas"].get(a, 0) for a in set(cur["cuentas"]) | set(old["cuentas"])})
            mv = [(a, v) for a, v in sorted(dif.items(), key=lambda t: -abs(t[1])) if abs(v) >= 0.05e6][:3]
            if mv:
                L.append("    - vs. mes anterior, por cuenta: " + "; ".join(f"{a.title()} {_amt(v)}" for a, v in mv))
        if cur["top"]:
            L.append("    - movimientos más grandes: " + "; ".join(
                f"{t[1]} {_amt(t[0])}" + (f" ({t[5]} pólizas)" if t[5] > 1 else f" ({t[3][8:10]}/{t[3][5:7]})")
                for t in cur["top"][:3]))
        shown += 1
        if shown >= k:
            break
    return L if shown else []


# ====================================================================================================
# resumen de la empresa (todo lo que sabe) y preguntas
# ====================================================================================================

MES_ES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
SUMMARY_LINES = [("Revenues", r"^revenues?$"), ("OPEX", r"^opex$"), ("Gross Profit", r"^gross profit$"),
                 ("Direct Line SG&A", r"^direct line sg&a$"), ("SG&A", r"^sg&a$"), ("EBITDA", r"^ebitda$"),
                 ("D&A", r"^d&a$"), ("Net Income", r"^net income$")]
SUMMARY_KPIS = [("total fleet", "Flota total", "int"), ("operational vehicles", "Flota operativa", "int"),
                ("oos", "Fuera de servicio (OOS)", "pct"), ("eop active drivers", "Drivers activos", "int"),
                ("headcount", "HC staff", "int"), ("supply hours", "Supply hours", "k"),
                ("revenue per hour", "RPH", "money1"), ("price per ride", "PPR", "money1"),
                ("trips per hour", "Trips por hora", "dec2"), ("financial uber take rate %", "Uber take rate", "pct"),
                ("maintenance per km", "Mantenimiento por km", "money2"), ("km travelled", "Km recorridos", "k"),
                ("hires", "Altas de drivers", "int"), ("exits", "Bajas de drivers", "int"), ("churn", "Churn", "pct"),
                ("performance ratio", "Performance ratio", "pct")]


def _m(v):
    return "n.d." if v is None else (f"(${abs(v) / 1e6:,.1f})" if v < 0 else f"${v / 1e6:,.1f}")


def _dm(v):
    return "n.d." if v is None else f"{'+' if v >= 0 else '-'}${abs(v) / 1e6:,.1f}"


def _kfmt(v, kind):
    if v is None:
        return "n.d."
    if kind == "pct":
        return f"{v * 100:.1f}%"
    if kind == "int":
        return f"{v:,.0f}"
    if kind == "k":
        return f"{v / 1000:,.1f}k"
    if kind == "money1":
        return f"${v:,.1f}"
    if kind == "money2":
        return f"${v:,.2f}"
    if kind == "dec2":
        return f"{v:.2f}"
    return f"{v:,.1f}"


def _first_row(rows, rx):
    return next((r for r in rows if re.search(rx, r["key"])), None)


def _ytd_budget(prep, row):
    """Budget acumulado del año hasta el mes (hoja Budget, misma etiqueta y orden)."""
    wb, ws, per = prep.get("wb"), prep.get("ws"), prep["period"]
    if wb is None or ws is None:
        return None
    tot, n = 0.0, 0
    for m in range(1, int(per[5:7]) + 1):
        ym = f"{per[:4]}-{m:02d}"
        if ym == per and row["b"] is not None:
            tot += row["b"]
            n += 1
            continue
        tab = _budget_lookup(wb, ws, ym)
        v = tab.get((row["key"], row["occ"])) if tab else None
        if v is not None:
            tot += v
            n += 1
    return tot if n else None


def load_news(company=None):
    """Hechos públicos (prensa) de datos/noticias_vemo.json: los de la empresa + los del grupo, más recientes primero."""
    try:
        with open(os.path.join(KNOWLEDGE_DIR, "noticias_vemo.json"), encoding="utf-8") as fh:
            hs = json.load(fh).get("hechos", [])
    except Exception:
        return []
    hs = [h for h in hs if company is None or h.get("empresa") in (company, "VEMO")]
    return sorted(hs, key=lambda h: (h.get("empresa") != company, h.get("fecha", "")), reverse=False)


def news_md(co, n_co=5, n_group=3):
    hs = load_news(co)
    own = sorted([h for h in hs if h["empresa"] == co], key=lambda h: h["fecha"], reverse=True)[:n_co]
    grp = sorted([h for h in hs if h["empresa"] == "VEMO"], key=lambda h: h["fecha"], reverse=True)[:n_group]
    if not own and not grp:
        return []
    L = ["", "## Contexto público (prensa; no sustituye tus cifras)", ""]
    L += [f"- {h['fecha']} · {h['hecho']} ([{h['fuente']}]({h['url']}))" for h in own]
    L += [f"- {h['fecha']} · Grupo: {h['hecho']} ([{h['fuente']}]({h['url']}))" for h in grp]
    return L


def load_highlights(company=None):
    """Tus Highlights & Lowlights publicados (todas las empresas). Marca 'stale' el tema copiado tal cual del mes anterior."""
    try:
        with open(os.path.join(KNOWLEDGE_DIR, "mfr_highlights.json"), encoding="utf-8") as fh:
            hl = json.load(fh)
    except Exception:
        return []
    hl.sort(key=lambda h: h["period"])
    seen = {}
    for h in hl:
        k = (h["company"], _plain(h["text"])[:200])
        h["stale"] = k in seen and seen[k] < h["period"]
        seen.setdefault(k, h["period"])
    return [h for h in hl if company is None or h["company"] == company]


def highlights_md(co, per, n_rec=6):
    hl = [h for h in load_highlights(co) if h["period"] <= per and h["kind"] in ("highlights", "lowlights")]
    if not hl:
        return []
    fresh = [h for h in hl if not h["stale"]]
    share = Counter(h["period"] for h in hl if h["stale"])
    tot = Counter(h["period"] for h in hl)
    ok = [pp for pp in tot if share[pp] < 0.5 * tot[pp]]  # un deck con la mayoría copiada del mes anterior no cuenta
    last = max(ok) if ok else max(tot)
    cur = [h for h in hl if h["period"] == last]
    mes = f"{MES_ES[int(last[5:7]) - 1]} {last[:4]}"
    tag = "de este mes" if last == per else f"último publicado: {mes}"
    L = ["", f"## Highlights & Lowlights de tus MFR ({tag})", ""]
    for kind, name in (("highlights", "Highlights"), ("lowlights", "Lowlights")):
        xs = [h for h in cur if h["kind"] == kind]
        if xs:
            L.append(f"**{name}**")
            L += [f"- {('**' + h['tema'] + ':** ') if h['tema'] else ''}{h['text'][:320]}" for h in xs]
            L.append("")
    stale_now = [h for h in hl if h["period"] == per and h["stale"]]
    if stale_now:
        L.append(f"_El deck de {MES_ES[int(per[5:7]) - 1]} todavía trae {len(stale_now)} de {tot[per]} temas copiados del mes anterior "
                 f"(aún no los actualizas)._")
        L.append("")
    pers = sorted({h["period"] for h in fresh})[-n_rec:]
    cnt = Counter()
    for h in fresh:
        if h["period"] in pers and h["tema"]:
            cnt[(h["kind"], re.sub(r"\s+", " ", h["tema"].strip()).title())] += 1
    rec = [(k, n) for k, n in cnt.most_common() if n >= 2]
    if rec:
        L.append(f"Temas que se repiten ({pers[0]} a {pers[-1]}): " + "; ".join(
            f"{t} ({'HL' if k == 'highlights' else 'LL'}, {n} meses)" for (k, t), n in rec[:8]) + ".")
    return L


def company_summary(prep, res, ctx, meetings, gold_path=None):
    """Panorama de la empresa en el mes: resultados, tendencia, operación, qué pasó, juntas y tus temas recurrentes."""
    co, per = prep["company"], prep["period"]
    rows, kpis = prep["data"]["rows"], prep["data"]["kpis"]
    y = per[:4]
    L = [f"# {co} — resumen de {MES_ES[int(per[5:7]) - 1]} {y}", ""]
    # 1. resultados del mes
    L += ["## Resultados del mes (MXN m)", "",
          "| Línea | Actual | Budget | vs. Budget | Mes anterior | MoM | YTD | YTD Budget | YTD vs. Budget |",
          "|---|---|---|---|---|---|---|---|---|"]
    rev = None
    for name, rx in SUMMARY_LINES:
        r = _first_row(rows, rx)
        if not r:
            continue
        ytd = sum(v or 0.0 for k, v in r["series"].items() if k.startswith(y))
        ytdb = _ytd_budget(prep, r)
        vsb = (r["m"] or 0) - r["b"] if r["b"] is not None else None
        mom = (r["m"] or 0) - (r["p"] or 0) if r["p"] is not None else None
        L.append(f"| {name} | {_m(r['m'])} | {_m(r['b'])} | {_dm(vsb)} | {_m(r['p'])} | {_dm(mom)} | {_m(ytd)} | {_m(ytdb)} | "
                 f"{_dm(ytd - ytdb) if ytdb is not None else 'n.d.'} |")
        if name == "Revenues":
            rev = r
    gp, eb = _first_row(rows, r"^gross profit$"), _first_row(rows, r"^ebitda$")
    if rev and rev["m"]:
        L.append("")
        L.append(f"Margen bruto {(gp['m'] or 0) / rev['m'] * 100:.1f}%" + (f" (budget {(gp['b'] or 0) / rev['b'] * 100:.1f}%)" if gp and gp["b"] is not None and rev["b"] else "")
                 + f" · margen EBITDA {(eb['m'] or 0) / rev['m'] * 100:.1f}%" + (f" (budget {(eb['b'] or 0) / rev['b'] * 100:.1f}%)" if eb and eb["b"] is not None and rev["b"] else "") + ".")
    # 2. tendencia
    L += ["", "## Tendencia (últimos 6 meses, MXN m)", ""]
    for name, rx in (("Revenues", r"^revenues?$"), ("Gross Profit", r"^gross profit$"), ("EBITDA", r"^ebitda$"), ("Net Income", r"^net income$")):
        r = _first_row(rows, rx)
        if r:
            ks = sorted(k for k in r["series"] if r["series"][k] is not None)[-6:]
            L.append(f"- **{name}:** " + " · ".join(f"{MES_ES[int(k[5:7]) - 1]} {(r['series'][k] or 0) / 1e6:,.1f}" for k in ks))
    try:
        L += history_md(co, per)
    except Exception:
        pass
    # 3. operación
    L += ["", "## Operación (KPIs)", "", "| KPI | Actual | Budget | Mes anterior |", "|---|---|---|---|"]
    for key, name, kind in SUMMARY_KPIS:
        k = kpis.get(key)
        if k and k["m"] is not None:
            L.append(f"| {name} | {_kfmt(k['m'], kind)} | {_kfmt(k['b'], kind) if k['b'] else 'n.d.'} | {_kfmt(k['p'], kind)} |")
    # 4. qué pasó en el mes (las notas de este mes de los renglones que más se movieron)
    moved = [x for x in ctx["lineas"] if any(ym == per for ym, _ in x["notas"])][:8]
    if moved:
        L += ["", "## Qué pasó en el mes (tus notas de este mes)", ""]
        for x in moved:
            t = next(t for ym, t in x["notas"] if ym == per)
            L.append(f"- **{x['label']}** (vs. Budget {fm(x['budget'])} · MoM {fm(x['mom'])}): {t[:260]}")
    # 5. juntas del cierre
    if meetings is not None and len(meetings):
        start, end = _window(per)
        facts = []
        for c in meetings.chunks:
            d = _date(c.get("date"))
            if c.get("kind") == "hecho" and d and start <= d <= end and co in (c.get("companies") or []):
                facts.append(c)
        facts.sort(key=lambda c: (-len(c.get("amounts") or []), c["date"]))
        if facts or ctx["pendientes"]:
            L += ["", "## Lo que se habló en juntas de este cierre (Granola)", ""]
            for c in facts[:8]:
                L.append(f"- {c['text'][:240]} _({c['meeting']}, {c['date']})_")
            for c in ctx["pendientes"][:6]:
                L.append(f"- Pendiente: {c['text'][:200]} _({c['meeting']}, {c['date']})_")
    # 6. contexto público
    try:
        L += news_md(co)
    except Exception:
        pass
    # 6b. tus Highlights & Lowlights
    try:
        L += highlights_md(co, per)
    except Exception:
        pass
    # 7. temas recurrentes en tus MFR
    if gold_path and os.path.exists(gold_path):
        try:
            with open(gold_path, encoding="utf-8") as fh:
                gold = [g for g in json.load(fh) if g["company"] == co and g["period"] < per]
            last = sorted({g["period"] for g in gold})[-4:]
            cnt = Counter()
            for g in gold:
                if g["period"] in last:
                    for m in learn_mentions(g["text"], g.get("box") or ""):
                        if m["key"] and m["key"] not in ("revenue", "opex", "gross profit", "ebitda", "direct line sg&a"):
                            cnt[m["key"]] += 1
            if cnt:
                L += ["", f"## Temas que más has explicado en tus MFR ({', '.join(last)})", "",
                      ", ".join(f"{_display(k)} ({n})" for k, n in cnt.most_common(8)) + "."]
        except Exception:
            pass
    return "\n".join(L) + "\n"


# ---------------------------------------------------------------- preguntas

ES_EN = [  # cómo preguntas en español -> renglón/KPI de la hoja
    (r"ingreso|revenue|ventas", "revenue"), (r"ride.?hailing|viajes", "ride hailing revenue"), (r"publicidad|advertising", "advertising revenue"),
    (r"opex|costo de operaci|gastos? operativ", "opex"), (r"utilidad bruta|gross profit|\bgp\b|margen bruto", "gross profit"),
    (r"ebitda", "ebitda"), (r"utilidad neta|net income|resultado neto", "net income"), (r"depreciaci|d&a", "d&a"),
    (r"take rate|comisi[oó]n de uber", "uber take rate"), (r"pago.*conductor|driver payment|n[oó]mina de conductor", "driver payments"),
    (r"mantenimiento|maintenance|taller", "maintenance"), (r"multa|penalt", "driving penalties & fines"),
    (r"seguro|insurance", "insurance"), (r"lavado|washing", "washing"), (r"gr[uú]a|towing", "towing service"),
    (r"arrendamiento|leasing", "vehicle leasing"), (r"telemetr|telematics", "telematics"), (r"energ[ií]a|ev fuel|electricidad", "ev fuel cost"),
    (r"chaas", "chaas"), (r"charging|cargador|infraestructura de carga", "charging infra"), (r"marketing|viniles|branding", "marketing"),
    (r"seguridad|security", "security"), (r"n[oó]mina|payroll", "payroll expense"), (r"honorarios|legal", "legal & professional fees"),
    (r"suscripci|licencia|subscription", "subscriptions"), (r"corporativ|cross.?functional|cf ", "cross-functional g&a"),
    (r"impuesto|taxes", "taxes"), (r"inter[eé]s", "interest expense"), (r"sg&a|gastos? de administraci", "sg&a"),
    (r"flota total|total fleet|flota\b", "total fleet"), (r"flota operativa|operational vehicles", "operational vehicles"),
    (r"drivers? activos|conductores activos|active drivers|\bdrivers\b|conductores", "eop active drivers"),
    (r"headcount|\bhc\b|plantilla", "headcount"), (r"supply hours|horas conectad|horas en l[ií]nea", "supply hours"),
    (r"\brph\b|revenue per hour|ingreso por hora", "revenue per hour"), (r"\bppr\b|precio por viaje|price per ride", "price per ride"),
    (r"churn|rotaci", "churn"), (r"\bkm\b|kil[oó]metros", "km travelled"), (r"fuera de servicio|\boos\b", "oos"),
]


def _q_company(q):
    for co in list(COMPANIES) + [c for c in HIST_COMPANIES if c not in COMPANIES]:
        if re.search(rf"\b{co}\b", q, re.I):
            return co
    return "DAE"


def _q_period(q, default):
    m = re.search(r"\b(20\d\d)[-/](\d{1,2})\b", q)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    words = _plain(q).split()
    for i, w in enumerate(words):
        if w in MONTH_NAMES:
            yr = next((x for x in words[i + 1:i + 3] if re.fullmatch(r"20\d\d", x)), None)
            try:
                return parse_periodo(f"{w} {yr}" if yr else w)
            except ValueError:
                pass
    return default


def _clauses(text):
    """Corta un box en sus drivers: por ';', '(a)…(h)', '(i)…' o 'partially offset by', sin cortar dentro de paréntesis."""
    body = text.split(":", 1)[1] if re.match(r"^\s*(Budget|MoM)\b", text) else text
    out, buf, depth, i = [], "", 0, 0
    while i < len(body):
        ch = body[i]
        if depth == 0:
            m = re.match(r"\(\s*(?:[a-h]|[ivx]+)\s*\)|partially offset by|This was", body[i:])
            if m or ch == ";":
                out.append(buf)
                buf = ""
                i += m.end() if m else 1
                continue
        depth += (ch == "(") - (ch == ")")
        depth = max(depth, 0)
        buf += ch
        i += 1
    out.append(buf)
    return [re.sub(r"^\s*(and|,)\s*", "", c).strip(" ,.;") for c in out if c.strip(" ,.;")]


HIST_Q = [(r"ingreso|revenue|ventas", ("nor", "revenue")), (r"utilidad bruta|gross profit|margen bruto|\bgp\b", ("gp",)),
          (r"margen financiero ajustado|adjusted financial", ("afm",)), (r"margen financiero|financial margin", ("fm",)),
          (r"ebitda", ("ebitda",)), (r"utilidad neta|net income|resultado neto|perdida", ("ni",)), (r"opex", ("opex",)),
          (r"sg&a|gastos de administraci", ("sga",)), (r"flota", ("fleet",)), (r"drivers|conductores", ("drivers",)),
          (r"supply hours|horas", ("sh",)), (r"cartera|portfolio", ("portfolio",)), (r"unidades", ("units",)),
          (r"originaci|colocaci", ("origination",))]


def _answer_hist(co, q, qp, H, year):
    keys = next((ks for rx, ks in HIST_Q if re.search(rx, qp)), None)
    if not keys:
        last = max((ym for y in H for k in ("nor", "revenue") for ym in ((H[y].get("actual") or {}).get("lines", {}).get(k) or {})),
                   default=None)
        return ("\n".join(history_md(co, last)) if last else "") or f"No tengo histórico de {co}."
    yrs = [year] if year else sorted(int(y) for y in H)
    out = [f"**{co} · histórico**"]
    for y in yrs:
        rec = H.get(str(y)) or {}
        a = (rec.get("actual") or {}).get("lines") or {}
        k = next((k for k in keys if k in a), None)
        if not k:
            out.append(f"- {y}: no lo encontré en el FY Financials de ese año.")
            continue
        b = ((rec.get("budget") or {}).get("lines") or {}).get(k)
        if b:  # budget de los mismos meses que el real
            b = {ym: v for ym, v in b.items() if ym in a[k]}
        av, bv = _hsum(a[k], k), (_hsum(b, k) if b else None)
        line = f"- {y} {HIST_NAMES[k]}: {_hfmt(av, k)} ({len(a[k])} meses)"
        if bv is not None:
            line += f" · budget {_hfmt(bv, k)}" + (f" ({_dm(av - bv)})" if k not in HIST_STOCK and k not in ("fleet", "drivers", "units", "origination", "sh") else "")
        line += f"  _(fuente: {rec['file'].split('/')[-1]} › {rec['actual']['sheet']}"
        line += ")_" if rec["actual"].get("cuadra") else "; ojo: esa hoja no cuadra Revenues + OPEX = GP, revísalo)_"
        out.append(line)
        if year:
            out.append("  " + " · ".join(f"{MES_ES[int(ym[5:7]) - 1]} {_hfmt(v, k)}" for ym, v in sorted(a[k].items())))
    return "\n".join(out)


def answer_question(root, q, meetings=None):
    co = _q_company(q)
    qp = _plain(q)
    if re.search(r"highlight|lowlight|\bhl\b|\bll\b|logro|lo bueno|lo malo|riesgo|preocupa", qp):  # no requiere el Excel
        per = _q_period(q, max([h["period"] for h in load_highlights(co)] or [_last_closed()]))
        hl = highlights_md(co, per)
        return "\n".join([f"**{co} · {MES_ES[int(per[5:7]) - 1]} {per[:4]}**"] +
                         (hl if hl else ["No tengo Highlights & Lowlights de esa empresa en tus decks."]))
    if re.search(r"noticia|prensa|news|publico|se dice de vemo", qp):
        nm = news_md(co, n_co=10, n_group=10)
        return "\n".join([f"**{co}**"] + (nm if nm else ["No tengo noticias guardadas."]))
    yq = re.search(r"\b(20[12]\d)\b", q)
    H = load_history().get(co) or {}
    hist_q = re.search(r"hist[oó]ric|historia|por a[nñ]o|anual|cada a[nñ]o|evoluci", qp)
    if H and (hist_q or co not in EJECUTA or (yq and int(yq.group(1)) < _dt.date.today().year)):
        return _answer_hist(co, q, qp, H, int(yq.group(1)) if yq else None)
    path = fy_file(root, co, None)
    if not path:
        return f"No encontré el FY Financials de {co}."
    try:
        per = _q_period(q, latest_period(path, co))
        load_style(perfil="calidad")
        prep = prepare(open_wb(path), co, per, path=path)
    except Exception as e:
        return f"No pude leer el P&L de {co} ({e}). Por ahora el P&L lo leo completo solo para {', '.join(EJECUTA)}; " \
               f"de {co} sí puedo responder sus Highlights & Lowlights."
    rows, kpis = prep["data"]["rows"], prep["data"]["kpis"]
    target = next((key for rx, key in ES_EN if re.search(rx, qp)), None)
    out = [f"**{co} · {MES_ES[int(per[5:7]) - 1]} {per[:4]}**"]
    rec = None
    if target:
        rec = next((r for r in rows if r["key"] == target), None) or kpis.get(target) or \
            next((r for r in rows if r["key"].startswith(target)), None)
    if rec:
        is_kpi = rec["key"] in kpis and rec is kpis.get(rec["key"])
        kind = next((k for key, _n, k in SUMMARY_KPIS if key == rec["key"]), "num")
        f = (lambda v: _kfmt(v, kind)) if is_kpi else _m
        out.append(f"- {rec['label']}: actual {f(rec['m'])} · budget {f(rec['b']) if rec['b'] is not None else 'n.d.'} · "
                   f"mes anterior {f(rec['p'])}  _(fuente: {prep['sheet']}!{rec['cell']})_")
        if not is_kpi:
            if rec["b"] is not None:
                out.append(f"- vs. Budget: {fm(delta(rec, 'budget'))} · MoM: {fm(delta(rec, 'mom'))}")
            else:
                out.append(f"- MoM: {fm(delta(rec, 'mom'))} (sin budget en el archivo)")
            if re.search(r"ytd|acumulad|en el a[nñ]o", qp):
                ytd = sum(v or 0.0 for k, v in rec["series"].items() if k.startswith(per[:4]))
                ytdb = _ytd_budget(prep, rec)
                out.append(f"- YTD: {_m(ytd)}" + (f" vs. budget {_m(ytdb)} ({_dm(ytd - ytdb)})" if ytdb is not None else ""))
            if re.search(r"por ?qu[eé]|why|explica|caus|driver|coment|raz[oó]n", qp):
                comp = compose(prep)["boxes"]
                box = next((x for x in comp if x.get("row") == rec["row"] and not x.get("missing")), None)
                if box:
                    out += ["", "Cómo lo redactaría en el MFR:", box["budget"]["text"], box["mom"]["text"]]
                else:  # renglón que no es box: la parte de los boxes que lo explica
                    words = {_plain(rec["label"]), _plain(_display(rec["key"]))}
                    cl = []
                    for x in comp:
                        for c in ("budget", "mom"):
                            if x.get("missing"):
                                continue
                            for seg in _clauses(x[c]["text"]):
                                if any(w and w in _plain(seg) for w in words):
                                    cl.append(f"- {x['box']} {'vs. Budget' if c == 'budget' else 'MoM'}: {seg.strip(' ,.;')}")
                    if cl:
                        out += ["", "Cómo lo redactaría en el MFR:"] + list(dict.fromkeys(cl))[:4]
            for ym, t in notes_timeline(rec, per, 3)[::-1] if rec.get("notes_hist") else []:
                out.append(f"- Nota de celda {ym}: {t[:300]}")
            D = load_dyn(co)
            if D:
                keys = _dyn_keys(D, rec["label"], _ancestors(prep, rec))
                cur = dyn_line(D, keys, per)
                if cur and cur["n"]:
                    out.append(f"- Dynamics {per}: {_m(cur['total'])} en {cur['n']} pólizas · cuentas: " + "; ".join(
                        f"{a.title()} {_m(v)}" for a, v in sorted(cur["cuentas"].items(), key=lambda t: -abs(t[1]))[:4]))
                    out += [f"    - {t[1]} {_amt(t[0])} ({t[2].title()}; {t[5]} póliza{'s' if t[5] > 1 else ''}, {t[3]})" for t in cur["top"][:4]]
    # juntas: solo fragmentos que hablan del tema (no de la empresa o el mes en general)
    stop = {"cuanto", "cuantos", "cuanta", "fue", "que", "por", "paso", "con", "del", "los", "las", "una", "hay", "mes",
            "como", "esta", "este", "subio", "bajo", "tiene", "tenemos", "cual", "donde", "ytd", "vs", "budget"}
    stop |= {_plain(c) for c in COMPANIES} | set(MONTH_NAMES)
    qterms = [w for w in re.findall(r"[a-z0-9&]+", qp) if len(w) > 2 and w not in stop and not re.fullmatch(r"20\d\d|\d+", w)]
    qwords = list(qterms)
    if target:
        qterms += (CTX_LINE_TERMS.get(target) or "").split()
    if meetings is not None and len(meetings) and qterms:
        hits = []
        for h in meetings.search(" ".join(qterms), k=20):
            ht = _plain(h["text"])
            if h.get("similarity", 0) >= 0.3 and (not h.get("companies") or co in h["companies"]) \
                    and any(re.search(rf"\b{re.escape(w)}", ht) for w in qterms if len(w) > 3):
                hits.append(h)
        if hits:
            out += ["", "Juntas:"] + [f"- {h['text'][:260]} _({h['meeting']}, {h['date']})_" for h in hits[:3]]
    # tus Highlights & Lowlights que tocan el tema
    ws = [w for w in qwords if len(w) > 3]
    if ws:
        hh = [h for h in load_highlights(co) if h["period"] <= per and not h["stale"]
              and sum(bool(re.search(rf"\b{re.escape(w)}", _plain(h["tema"] + " " + h["text"]))) for w in ws) >= min(2, len(ws))]
        hh.sort(key=lambda h: h["period"], reverse=True)
        if hh:
            out += ["", "En tus Highlights & Lowlights:"]
            tag = {"highlights": "HL", "lowlights": "LL", "comentario_pl": "Comentario P&L"}
            out += [f"- {h['period']} {tag.get(h['kind'], h['kind'])} · {h['tema']}: {h['text'][:260]}" for h in hh[:3]]
    # prensa que toca el tema
    if ws:
        nn = [h for h in load_news(co) if sum(bool(re.search(rf"\b{re.escape(w)}", _plain(h["tema"] + " " + h["hecho"]))) for w in ws) >= min(2, len(ws))]
        if nn:
            out += ["", "En prensa:"] + [f"- {h['fecha']} · {h['hecho']} ({h['fuente']})" for h in nn[:2]]
    # lo que tú escribiste en tus MFR sobre ese renglón
    gold_path = os.path.join(KNOWLEDGE_DIR, "mfr_gold.json")
    if target and os.path.exists(gold_path):
        with open(gold_path, encoding="utf-8") as fh:
            gold = [g for g in json.load(fh) if g["company"] == co and g["period"] <= per]
        gold.sort(key=lambda g: g["period"])
        seen, said = set(), []
        for g in gold:  # un box que repite el texto del mes anterior es un box sin actualizar: se ignora
            if g["text"] in seen:
                continue
            seen.add(g["text"])
            if _plain(g.get("box") or "") == target:
                said.append((g["period"], g["comp"], g["text"][:260]))
                continue
            for m in learn_mentions(g["text"], g.get("box") or ""):
                if m["key"] == target:
                    amt = m["paren"].split(")", 1)[0]
                    seg = next((c for c in _clauses(g["text"]) if amt and amt in c), None) or \
                        next((c for c in _clauses(g["text"]) if m["phrase"] and _plain(m["phrase"]) in _plain(c)), None)
                    said.append((g["period"], g["comp"], seg or f"{m['phrase']} {m['paren']}".strip()))
                    break
        if said:
            out += ["", "Lo que escribiste en tus MFR:"]
            out += [f"- {pp} ({'vs. Budget' if c == 'budget' else 'MoM'}): {t}" for pp, c, t in sorted(said, reverse=True)[:3]]
    if len(out) == 1:
        return out[0] + "\nNo tengo evidencia en tus archivos para responder eso (no invento)."
    return "\n".join(out)


# ====================================================================================================
# menú y comandos
# ====================================================================================================

EJECUTA = ("DAE",)  # empresas con motor MFR activo (el vocabulario se aprende de todas)
DATA_STYLE = "mfr_style.json"


def fy_file(root, company, year=None):
    cands = candidate_files(root, company, COMPANY_FILE_KEYWORDS, EXCLUDE_KEYWORDS, year=year)
    if not cands:
        scan_tree(root, fresh=True)
        cands = candidate_files(root, company, COMPANY_FILE_KEYWORDS, EXCLUDE_KEYWORDS, year=year)
    return next((c for c in cands if "/0. vf/" in c.replace("\\", "/").lower()), cands[0] if cands else None)


def get_meetings():
    try:
        return MeetingMemory.load()
    except Exception as e:  # una junta mal formada nunca debe tumbar los comentarios
        print(f"[juntas] no pude leer la carpeta de juntas ({e}); sigo sin ese contexto.")
        return None


def _style_path():
    return os.path.join(MODELS_DIR, DATA_STYLE)


def _decks_changed(root):
    try:
        last = os.path.getmtime(_style_path()) if os.path.exists(_style_path()) else 0
        decks = pick_decks(os.path.join(root, "2. Monthly"))
        return any(os.path.getmtime(p) > last for p in decks.values())
    except Exception:
        return False


def comentar_empresa(root, company, periodo, outdir, estilo="calidad", meetings=None):
    year = int(periodo[:4]) if periodo else None
    path = fy_file(root, company, year)
    if not path:
        raise RuntimeError(f"no encontré el FY Financials de {company}")
    load_style(perfil=estilo)
    per = periodo or latest_period(path, company)
    prep = prepare(open_wb(path), company, per, path=path)
    res = compose(prep)
    ctx = build_context(res, prep, meetings)
    try:
        dyn_build(root, company, verbose=False)
        ctx["dynamics_md"] = dyn_md(prep, ctx, company)
        D = load_dyn(company)
        if D and prep["period"] in D["meses"]:  # Board vs. Dynamics: diferencias grandes antes de explicar
            for x in ctx["lineas"][:14]:
                n = x["node"]
                cur = dyn_line(D, _dyn_keys(D, n["label"], _ancestors(prep, n)), prep["period"])
                if cur and cur["n"] and abs((n["m"] or 0) - cur["total"]) >= 0.3e6:
                    res["flags"].append(f"{x['label']}: el Board {_m(n['m'])} y Dynamics {_m(cur['total'])} difieren por "
                                        f"{_amt((n['m'] or 0) - cur['total'])} (provisión, reclasificación o ajuste manual); "
                                        f"confírmalo antes de explicar la variación.")
        if D and prep["period"] in D["meses"]:  # renglón sin nota: lo que muestra Dynamics como pista
            for i, f in enumerate(res["flags"]):
                if "sin nota" not in f:
                    continue
                lab = _plain(f.split(" (")[0].split(":")[0])
                n = next((r for r in prep["data"]["rows"] if _plain(r["label"]) == lab), None)
                cur = dyn_line(D, _dyn_keys(D, n["label"], _ancestors(prep, n)), prep["period"]) if n else None
                if cur and cur["top"]:
                    res["flags"][i] = f + " Dynamics muestra: " + "; ".join(f"{t[1]} {_amt(t[0])}" for t in cur["top"][:2]) + "."
    except Exception as e:
        ctx["dynamics_md"] = [f"", f"_(no pude leer Dynamics: {e})_"]
    for j in ctx["juntas"]:  # una junta que dice la MISMA cifra que una variación: revisar si es la causa
        for h in j["hits"]:
            if h.get("match"):
                res["flags"].append(f"{j['label']}: la junta «{h['meeting']}» ({h['date']}) menciona {fm(h['match'][0])}, "
                                    f"igual a la variación del Excel; revisa si es la causa: “{h['text'][:160]}”")
    try:
        summary = company_summary(prep, res, ctx, meetings, os.path.join(KNOWLEDGE_DIR, "mfr_gold.json"))
    except Exception as e:  # el resumen nunca debe impedir los comentarios
        summary = f"# {company} — resumen {per}\n\n(no pude armar el resumen: {e})\n"
    res["summary"] = summary
    res["_prep"], res["_ctx"] = prep, ctx  # para el dashboard
    md = summary + "\n---\n\n" + render(res, show_evidence=False) + "\n" + render_context(ctx, per)
    os.makedirs(outdir, exist_ok=True)
    base = os.path.join(outdir, f"{company}_{per}")
    with open(base + ".md", "w", encoding="utf-8") as fh:
        fh.write(md)
    return res, base


def learn_mfr(root, company="DAE", dry=False, verbose=True):
    """Lee TUS boxes publicados en los MFR (todas las empresas) y ajusta el estilo para `company`.
    Solo adopta el estilo nuevo si se parece a ti al menos tanto como el anterior."""
    monthly = os.path.join(root, "2. Monthly")
    gold_path = os.path.join(KNOWLEDGE_DIR, "mfr_gold.json")
    if os.path.isdir(monthly):
        os.environ.setdefault("MFR_DESDE", "2023-01")
        gold_update(monthly, gold_path, verbose=verbose)
    try:
        dyn_build(root, company, verbose=verbose)
    except Exception as e:
        print(f"  (no pude leer Dynamics: {e})")
    try:
        print("  Histórico de FY Financials (todas las empresas y años) ...")
        build_history(root, verbose=verbose)
    except Exception as e:
        print(f"  (no pude actualizar el histórico: {e})")
    if not os.path.exists(gold_path):
        print("  (no encontré decks MFR con boxes de P&L; nada que aprender todavía)")
        return None
    fy = fy_file(root, company, _dt.date.today().year)
    if not fy:
        print(f"  no encontré el FY Financials de {company}")
        return None
    prev = None
    if os.path.exists(_style_path()):
        try:
            with open(_style_path(), encoding="utf-8") as fh:
                prev = json.load(fh)
        except Exception:
            prev = None
    tmp = _style_path() + ".nuevo"
    res = learn_run(company, fy, gold_path, out_path=tmp,
                    report_path=os.path.join(OUTPUT_DIR_DEFAULT, f"mfr_backtest_{company}.md"), verbose=True)
    if prev and prev.get("style"):
        old_score, _ = score_style(prev["style"], company, fy, gold_path, res["trained_on"])
        if old_score > res["score_after"] + 0.005:
            print(f"  Estilo nuevo {res['score_after']:.0%} < anterior {old_score:.0%}: me quedo con el anterior.")
            os.remove(tmp)
            return res
    if dry:
        os.remove(tmp)
        return res
    os.replace(tmp, _style_path())
    print(f"  Estilo guardado. Reporte: salidas/mfr_backtest_{company}.md")
    try:
        q = quality_score(company, fy, os.path.join(KNOWLEDGE_DIR, "oro_calidad.json"))
        if q:
            print(f"  Calidad vs. boxes de referencia: drivers {q['drivers']:.0%} ({', '.join(q['periodos'])}).")
    except Exception as e:
        print(f"  (no pude calificar contra los boxes de referencia: {e})")
    return res


def _root_or_ask():
    root = resolve_root(None)
    while not root:
        print("No encuentro tu carpeta de FP&A en OneDrive.")
        cand = input("Pega la ruta de la carpeta '... - FP&A': ").strip().strip('"')
        if cand and os.path.isdir(cand):
            cfg = load_config()
            cfg["root"] = cand
            save_config(cfg)
            root = cand
    return root


def run_comentar(root, company, periodo, estilo="calidad", aprender="auto"):
    if aprender != "no" and (aprender == "si" or _decks_changed(root)):
        print("[aprendizaje] hay decks MFR nuevos o editados: ajusto el estilo antes de redactar ...")
        try:
            learn_mfr(root, company=company, verbose=False)
        except Exception as e:
            print(f"[aprendizaje] omitido ({e}); sigo con el estilo actual.")
    if not os.path.exists(os.path.join(KNOWLEDGE_DIR, "historico.json")):
        print("[histórico] leo una vez los FY Financials de todos los años (tarda unos minutos) ...")
        try:
            build_history(root, verbose=False)
        except Exception as e:
            print(f"[histórico] omitido ({e})")
    print(f"\nGenerando {company} {periodo or '(mes vigente)'} ...")
    res, base = comentar_empresa(root, company, periodo, OUTPUT_DIR_DEFAULT, estilo=estilo, meetings=get_meetings())
    print("=" * 78)
    print(_console(res.get("summary", "")))
    print("=" * 78)
    print("COMENTARIOS MFR (listos para pegar)\n")
    print(paste_text(res))
    if res["flags"]:
        print("Revisar antes de pegar:")
        for i, f in enumerate(res["flags"], 1):
            print(f"  {i}. {f}")
    print(f"\n[guardado] {base}.md")


def _console(md):
    """Markdown -> texto legible en la consola (sin asteriscos ni guiones bajos de formato)."""
    out = []
    for ln in md.splitlines():
        if re.fullmatch(r"\|[-|\s]+\|", ln.strip()):
            continue
        ln = re.sub(r"^#+\s*", "", ln)
        ln = ln.replace("**", "")
        ln = re.sub(r"(?<!\w)_(.+?)_(?!\w)", r"\1", ln)
        if ln.startswith("|"):
            cs = [c.strip() for c in ln.strip("|").split("|")]
            ln = "  " + cs[0].ljust(26) + "  ".join(c.rjust(11) for c in cs[1:])
        out.append(ln)
    return "\n".join(out)


def run_preguntar(root, meetings=None):
    meetings = meetings if meetings is not None else get_meetings()
    print("Pregunta lo que quieras de la empresa (ej. '¿cuánto fue el EBITDA de DAE en septiembre?', "
          "'¿por qué subió mantenimiento?'). Enter vacío = regresar.")
    while True:
        q = input("\n? ").strip()
        if not q:
            return
        try:
            print(_console(answer_question(root, q, meetings)))
        except Exception as e:
            print(f"!! No pude responder: {e}")


def menu():
    root = _root_or_ask()
    while True:
        print("\n=== VEMO IA — resumen y boxes del MFR ===")
        print("  1) Resumen de la empresa + comentarios MFR")
        print("  2) Preguntar")
        print("  3) Aprender de mis MFR publicados (después de cerrar cada mes)")
        print("  4) Abrir dashboard en el navegador")
        print("  5) Dynamics (API): configurar / probar / actualizar")
        print("  6) Salir")
        op = input("Opción [1]: ").strip() or "1"
        if op == "6":
            return
        if op == "5":
            sub = input("  a) configurar  b) probar conexión  c) actualizar ahora [c]: ").strip().lower() or "c"
            if sub == "a":
                dyn_api_configurar("DAE")
            elif sub == "b":
                dyn_api_probar("DAE")
            else:
                dyn_build(root, "DAE", verbose=True, force=True)
            continue
        if op == "4":
            run_dashboard(root)
            continue
        if op == "3":
            learn_mfr(root, company="DAE")
            continue
        if op == "2":
            run_preguntar(root)
            continue
        emp = (input(f"Empresa ({'/'.join(COMPANIES)}) [DAE]: ").strip() or "DAE").upper()
        if emp not in EJECUTA:
            print(f"  Por ahora el motor MFR corre solo para {', '.join(EJECUTA)}.")
            continue
        txt = input("Mes (ej. 2026-09, 'septiembre 2026'; Enter = el mes vigente del archivo): ").strip()
        try:
            periodo = parse_periodo(txt) if txt else None
        except ValueError:
            print("  No entendí el mes. Ejemplos: 2026-09, sep-26, septiembre 2026.")
            continue
        try:
            run_comentar(root, emp, periodo)
        except Exception as e:
            print(f"!! No pude generar {emp} {periodo or ''}: {e}")


# ====================================================================================================
# dashboard local (http://127.0.0.1:8765): eliges empresa y mes, y genera resumen + comentarios en vivo
# ====================================================================================================

DASH_PORT = 8765


def _dash_trend(prep, n=12):
    rows, per = prep["data"]["rows"], prep["period"]
    out, labels = {}, None
    for name, rx in (("Revenues", r"^revenues?$"), ("Gross Profit", r"^gross profit$"), ("EBITDA", r"^ebitda$"),
                     ("Net Income", r"^net income$")):
        r = _first_row(rows, rx)
        if not r:
            continue
        ks = sorted(k for k, v in r["series"].items() if v is not None and k <= per)[-n:]
        labels = labels or ks
        out[name] = [round((r["series"].get(k) or 0) / 1e6, 2) for k in labels]
    return {"labels": [f"{MES_ES[int(k[5:7]) - 1]}-{k[2:4]}" for k in (labels or [])], "series": out}


def dash_generar(root, company, periodo, meetings):
    company = company.upper()
    if company in EJECUTA:
        res, base = comentar_empresa(root, company, periodo, OUTPUT_DIR_DEFAULT, meetings=meetings)
        prep, ctx = res.pop("_prep"), res.pop("_ctx")
        boxes = [{"box": b["box"], "budget": b["budget"]["text"], "mom": b["mom"]["text"]}
                 for b in res["boxes"] if not b.get("missing")]
        return {"empresa": company, "periodo": prep["period"], "motor": True, "resumen": res.get("summary", ""),
                "boxes": boxes, "pegar": paste_text(res), "flags": res["flags"],
                "contexto": render_context(ctx, prep["period"]), "trend": _dash_trend(prep),
                "archivo": os.path.basename(base) + ".md", "fuente": f"{os.path.basename(prep['file'])} › {prep['sheet']}"}
    per = periodo or _last_closed()
    L = [f"# {company} — {MES_ES[int(per[5:7]) - 1]} {per[:4]}", ""]
    for part in (history_md, ):
        try:
            L += part(company, per)
        except Exception:
            pass
    for fn in (lambda: news_md(company), lambda: highlights_md(company, per)):
        try:
            L += fn()
        except Exception:
            pass
    return {"empresa": company, "periodo": per, "motor": False, "resumen": "\n".join(L), "boxes": [], "pegar": "",
            "flags": [], "contexto": "", "trend": None,
            "aviso": f"El motor de comentarios MFR por ahora corre para {', '.join(EJECUTA)}. De {company} te muestro "
                     f"histórico, prensa y tus Highlights & Lowlights."}


DASH_HTML = r"""<!doctype html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>VEMO IA · MFR</title>
<style>
:root{--bg:#f6f7f9;--card:#fff;--ink:#16181d;--mute:#667085;--line:#e4e7ec;--acc:#0b7a5c;--acc2:#e7f5f0;--warn:#b54708;--warnbg:#fff6ed;
--s1:#0b7a5c;--s2:#2e6fd8;--s3:#c4320a;--s4:#7a5af8}
@media (prefers-color-scheme:dark){:root{--bg:#0f1115;--card:#171a21;--ink:#e6e8ec;--mute:#98a2b3;--line:#2a2f3a;--acc:#3ccf9e;--acc2:#12312a;
--warn:#fdb022;--warnbg:#2b2112;--s1:#3ccf9e;--s2:#6aa0ff;--s3:#ff8a5c;--s4:#a48afb}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.5 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
header{display:flex;align-items:center;gap:16px;padding:14px 24px;border-bottom:1px solid var(--line);background:var(--card);position:sticky;top:0;z-index:5;flex-wrap:wrap}
h1{font-size:17px;margin:0 12px 0 0;letter-spacing:-.01em}h1 b{color:var(--acc)}
.seg{display:flex;border:1px solid var(--line);border-radius:8px;overflow-x:auto;max-width:100%}
.seg button{white-space:nowrap;border:0;background:transparent;color:var(--ink);padding:7px 13px;cursor:pointer;font:inherit;border-right:1px solid var(--line)}
.seg button:last-child{border-right:0}.seg button.on{background:var(--acc);color:#fff}
.seg button small{opacity:.6;margin-left:4px;font-size:10px}
input[type=month],input[type=text]{font:inherit;padding:6px 10px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink)}
.btn{font:inherit;font-weight:600;padding:7px 16px;border-radius:8px;border:0;background:var(--acc);color:#fff;cursor:pointer}
.btn.ghost{background:transparent;color:var(--acc);border:1px solid var(--acc)}.btn:disabled{opacity:.5;cursor:wait}
#status{color:var(--mute);font-size:13px;margin-left:auto}
main{max-width:1180px;margin:0 auto;padding:20px 24px 60px}
.tabs{display:flex;gap:4px;border-bottom:1px solid var(--line);margin-bottom:16px;flex-wrap:wrap}
.tabs button{border:0;background:none;color:var(--mute);padding:9px 14px;font:inherit;cursor:pointer;border-bottom:2px solid transparent}
.tabs button.on{color:var(--ink);border-bottom-color:var(--acc);font-weight:600}
.tabs .n{background:var(--warnbg);color:var(--warn);border-radius:10px;padding:0 7px;font-size:11px;margin-left:5px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px 18px;margin-bottom:14px}
.box h3{margin:0 0 10px;font-size:15px;display:flex;justify-content:space-between;align-items:center}
.cmt{position:relative;padding:10px 12px 10px 14px;border-left:3px solid var(--acc);background:var(--acc2);border-radius:6px;margin:8px 0;white-space:pre-wrap}
.cmt .lbl{font-size:11px;font-weight:700;color:var(--acc);text-transform:uppercase;letter-spacing:.04em;display:block;margin-bottom:2px}
.copy{font:12px system-ui;border:1px solid var(--line);background:var(--card);color:var(--mute);border-radius:6px;padding:3px 9px;cursor:pointer}
.copy.ok{color:var(--acc);border-color:var(--acc)}
.flag{padding:9px 12px;background:var(--warnbg);border-radius:8px;margin:6px 0;color:var(--ink)}
.flag b{color:var(--warn)}
.aviso{padding:10px 14px;border-radius:8px;background:var(--warnbg);margin-bottom:14px}
.md table{border-collapse:collapse;width:100%;margin:8px 0 14px;font-variant-numeric:tabular-nums;font-size:13px}
.md th,.md td{padding:6px 9px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap}
.md th:first-child,.md td:first-child{text-align:left}.md th{color:var(--mute);font-weight:600;font-size:12px}
.md .tw{overflow-x:auto}.md h1{font-size:19px;margin:4px 0 10px}.md h2{font-size:15px;margin:22px 0 8px;padding-top:12px;border-top:1px solid var(--line)}
.md ul{padding-left:20px;margin:6px 0}.md li{margin:3px 0}.md a{color:var(--s2)}.md em{color:var(--mute)}
.chart{margin-bottom:6px}.legend{display:flex;gap:14px;font-size:12px;color:var(--mute);flex-wrap:wrap}.legend i{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:5px;vertical-align:-1px}
svg text{fill:var(--mute);font-size:11px}
.empty{color:var(--mute);text-align:center;padding:60px 0}
.qa{display:flex;gap:8px}.qa input{flex:1}
.spin{display:inline-block;width:12px;height:12px;border:2px solid var(--line);border-top-color:var(--acc);border-radius:50%;animation:r 0.8s linear infinite;vertical-align:-2px;margin-right:6px}
@keyframes r{to{transform:rotate(360deg)}}
.foot{color:var(--mute);font-size:12px;margin-top:8px}
@media (max-width:640px){header{padding:12px 16px}main{padding:16px}}
</style></head><body>
<header>
  <h1><b>VEMO IA</b> · MFR</h1>
  <div class="seg" id="emp"></div>
  <input type="month" id="mes">
  <button class="btn" id="go">Generar</button>
  <button class="btn ghost" id="dyn" title="Baja de nuevo las pólizas de Dynamics">Actualizar Dynamics</button>
  <span id="status">Elige empresa y mes.</span>
</header>
<main>
  <div class="tabs" id="tabs">
    <button data-t="cmt" class="on">Comentarios MFR</button>
    <button data-t="res">Resumen de la empresa</button>
    <button data-t="rev">Revisar<span class="n" id="nflags" hidden></span></button>
    <button data-t="ctx">Contexto (notas y juntas)</button>
    <button data-t="ask">Preguntar</button>
  </div>
  <section id="t-cmt"><div class="empty">Genera para ver los boxes listos para pegar.</div></section>
  <section id="t-res" hidden></section>
  <section id="t-rev" hidden></section>
  <section id="t-ctx" hidden></section>
  <section id="t-ask" hidden>
    <div class="card"><div class="qa"><input type="text" id="q" placeholder="Ej. ¿por qué subió mantenimiento? · EBITDA de DAE 2024 · highlights de LTO · noticias de VCN">
    <button class="btn" id="qgo">Preguntar</button></div><div id="qa-out" class="md" style="margin-top:12px"></div></div>
  </section>
</main>
<script>
const $=s=>document.querySelector(s);let EMP="DAE",INFO=null,timer=null;
const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
function inline(s){s=esc(s);return s.replace(/\[([^\]]+)\]\((https?:[^)\s]+)\)/g,'<a href="$2" target="_blank" rel="noopener">$1</a>')
 .replace(/\*\*(.+?)\*\*/g,"<b>$1</b>").replace(/(^|[\s(])_(.+?)_(?=[\s).,;:]|$)/g,"$1<em>$2</em>");}
function md(src){const L=String(src||"").split("\n");let h="",i=0;
 while(i<L.length){const l=L[i];
  if(/^\s*\|/.test(l)){const rows=[];while(i<L.length&&/^\s*\|/.test(L[i])){rows.push(L[i]);i++}
   const cells=r=>r.trim().replace(/^\||\|$/g,"").split("|").map(c=>c.trim());
   const body=rows.filter(r=>!/^\s*\|[\s|:-]+\|\s*$/.test(r));
   h+='<div class="tw"><table><thead><tr>'+cells(body[0]).map(c=>"<th>"+inline(c)+"</th>").join("")+"</tr></thead><tbody>"+
     body.slice(1).map(r=>"<tr>"+cells(r).map(c=>"<td>"+inline(c)+"</td>").join("")+"</tr>").join("")+"</tbody></table></div>";continue}
  if(/^\s*-\s+/.test(l)){h+="<ul>";while(i<L.length&&/^\s*-\s+/.test(L[i])){h+="<li>"+inline(L[i].replace(/^\s*-\s+/,""))+"</li>";i++}h+="</ul>";continue}
  if(/^#\s/.test(l))h+="<h1>"+inline(l.slice(2))+"</h1>";else if(/^##+\s/.test(l))h+="<h2>"+inline(l.replace(/^#+\s/,""))+"</h2>";
  else if(/^---+$/.test(l.trim()))h+="";else if(l.trim())h+="<p>"+inline(l)+"</p>";i++}
 return h}
function tab(t){document.querySelectorAll("#tabs button").forEach(b=>b.classList.toggle("on",b.dataset.t===t));
 ["cmt","res","rev","ctx","ask"].forEach(x=>$("#t-"+x).hidden=x!==t)}
document.querySelectorAll("#tabs button").forEach(b=>b.onclick=()=>tab(b.dataset.t));
async function copy(txt,btn){try{await navigator.clipboard.writeText(txt)}catch(e){const a=document.createElement("textarea");a.value=txt;document.body.appendChild(a);a.select();document.execCommand("copy");a.remove()}
 btn.classList.add("ok");const o=btn.textContent;btn.textContent="Copiado";setTimeout(()=>{btn.classList.remove("ok");btn.textContent=o},1400)}
function chart(tr){if(!tr||!tr.labels||!tr.labels.length)return"";const W=1080,H=220,P={l:44,r:28,t:12,b:26};
 const cols={"Revenues":"var(--s1)","Gross Profit":"var(--s2)","EBITDA":"var(--s3)","Net Income":"var(--s4)"};
 const all=Object.values(tr.series).flat();let mn=Math.min(0,...all),mx=Math.max(0,...all);const pad=(mx-mn)*.08||1;mn-=pad;mx+=pad;
 const x=i=>P.l+i*(W-P.l-P.r)/Math.max(1,tr.labels.length-1),y=v=>P.t+(mx-v)*(H-P.t-P.b)/(mx-mn);
 let g="";for(let k=0;k<=4;k++){const v=mn+(mx-mn)*k/4;g+=`<line x1="${P.l}" x2="${W-P.r}" y1="${y(v)}" y2="${y(v)}" stroke="var(--line)"/><text x="${P.l-6}" y="${y(v)+4}" text-anchor="end">${v.toFixed(0)}</text>`}
 g+=`<line x1="${P.l}" x2="${W-P.r}" y1="${y(0)}" y2="${y(0)}" stroke="var(--mute)" stroke-width=".6"/>`;
 tr.labels.forEach((l,i)=>g+=`<text x="${x(i)}" y="${H-8}" text-anchor="middle">${esc(l)}</text>`);
 for(const [n,vs] of Object.entries(tr.series)){g+=`<polyline fill="none" stroke="${cols[n]}" stroke-width="2.2" points="${vs.map((v,i)=>x(i)+","+y(v)).join(" ")}"/>`+
  vs.map((v,i)=>`<circle cx="${x(i)}" cy="${y(v)}" r="2.8" fill="${cols[n]}"><title>${esc(n)} ${esc(tr.labels[i])}: ${v.toFixed(1)}m</title></circle>`).join("")}
 return `<div class="card chart"><div class="legend">${Object.keys(tr.series).map(n=>`<span><i style="background:${cols[n]}"></i>${esc(n)}</span>`).join("")}<span>MXN m · últimos 12 meses</span></div>
 <svg viewBox="0 0 ${W} ${H}" width="100%" role="img" aria-label="Tendencia">${g}</svg></div>`}
function render(d){
 const av=d.aviso?`<div class="aviso">${esc(d.aviso)}</div>`:"";
 if(d.boxes.length){let h=`<div class="card" style="display:flex;justify-content:space-between;align-items:center"><div><b>${esc(d.empresa)} · ${esc(d.periodo)}</b><div class="foot">Fuente: ${esc(d.fuente||"")} · guardado en salidas/${esc(d.archivo||"")}</div></div><button class="copy" id="cpall">Copiar todo</button></div>`;
  d.boxes.forEach((b,i)=>{h+=`<div class="card box"><h3>${esc(d.empresa)} – ${esc(b.box)}</h3>
   <div class="cmt"><span class="lbl">vs. Budget</span>${esc(b.budget)}</div><button class="copy" data-c="${i}b">Copiar</button>
   <div class="cmt"><span class="lbl">MoM</span>${esc(b.mom)}</div><button class="copy" data-c="${i}m">Copiar</button></div>`});
  $("#t-cmt").innerHTML=h;$("#cpall").onclick=e=>copy(d.pegar,e.target);
  document.querySelectorAll("[data-c]").forEach(b=>b.onclick=()=>{const k=b.dataset.c,bx=d.boxes[+k.slice(0,-1)];copy(k.endsWith("b")?bx.budget:bx.mom,b)})}
 else $("#t-cmt").innerHTML=av+`<div class="empty">Sin comentarios MFR para ${esc(d.empresa)} todavía. Revisa la pestaña Resumen.</div>`;
 $("#t-res").innerHTML=av+chart(d.trend)+`<div class="card md">${md(d.resumen)}</div>`;
 $("#t-rev").innerHTML=d.flags.length?`<div class="card">${d.flags.map((f,i)=>`<div class="flag"><b>${i+1}.</b> ${esc(f)}</div>`).join("")}</div>`:`<div class="empty">Nada pendiente de revisar.</div>`;
 $("#nflags").hidden=!d.flags.length;$("#nflags").textContent=d.flags.length;
 $("#t-ctx").innerHTML=d.contexto?`<div class="card md">${md(d.contexto)}</div>`:`<div class="empty">Sin contexto adicional.</div>`}
async function generar(){const per=$("#mes").value;$("#go").disabled=true;const t0=Date.now();
 timer=setInterval(()=>$("#status").innerHTML=`<span class="spin"></span>Generando ${EMP} ${per}… ${((Date.now()-t0)/1000).toFixed(0)} s`,250);
 try{const r=await fetch(`/api/generar?empresa=${EMP}&periodo=${per}`);const d=await r.json();if(d.error)throw new Error(d.error);
  render(d);$("#status").textContent=`${d.empresa} ${d.periodo} listo en ${((Date.now()-t0)/1000).toFixed(0)} s`;tab(d.boxes.length?"cmt":"res")}
 catch(e){$("#status").textContent="Error: "+e.message}finally{clearInterval(timer);$("#go").disabled=false}}
async function preguntar(){const q=$("#q").value.trim();if(!q)return;$("#qgo").disabled=true;$("#qa-out").innerHTML='<span class="spin"></span>Buscando…';
 try{const r=await fetch(`/api/preguntar?q=${encodeURIComponent(q)}`);const d=await r.json();$("#qa-out").innerHTML=d.error?esc("Error: "+d.error):md(d.respuesta)}
 finally{$("#qgo").disabled=false}}
$("#dyn").onclick=async()=>{$("#dyn").disabled=true;$("#status").innerHTML='<span class="spin"></span>Actualizando Dynamics…';
 try{const d=await(await fetch("/api/dynamics")).json();$("#status").textContent=d.error?("Error: "+d.error):
  (d.aviso?d.aviso:`Dynamics (${d.fuente||"sin fuente"}) actualizado ${d.actualizado||""} · último movimiento ${d.ultimo_mov||"n.d."}`)}
 finally{$("#dyn").disabled=false}};
$("#go").onclick=generar;$("#qgo").onclick=preguntar;$("#q").onkeydown=e=>{if(e.key==="Enter")preguntar()};
fetch("/api/info").then(r=>r.json()).then(d=>{INFO=d;$("#mes").value=d.ultimo;
 $("#emp").innerHTML=d.empresas.map(e=>`<button data-e="${e.code}" class="${e.code===EMP?"on":""}" title="${e.motor?"Comentarios MFR + resumen":"Resumen, histórico y HL/LL"}">${e.code}${e.motor?"":"<small>resumen</small>"}</button>`).join("");
 document.querySelectorAll("#emp button").forEach(b=>b.onclick=()=>{EMP=b.dataset.e;document.querySelectorAll("#emp button").forEach(x=>x.classList.toggle("on",x===b))})});
</script></body></html>"""


def run_dashboard(root, port=DASH_PORT, open_browser=True):
    import http.server
    import threading
    import urllib.parse
    import webbrowser
    meetings = get_meetings()
    lock = threading.Lock()
    empresas = [{"code": c, "motor": c in EJECUTA} for c in list(COMPANIES) + [c for c in ("CORP", "ZEE") if c not in COMPANIES]]

    class Handler(http.server.BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def _send(self, code, body, ctype="application/json; charset=utf-8"):
            data = body.encode("utf-8") if isinstance(body, str) else body
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(data)

        def _json(self, obj, code=200):
            self._send(code, json.dumps(obj, ensure_ascii=False, default=str))

        def do_GET(self):
            u = urllib.parse.urlparse(self.path)
            q = {k: v[0] for k, v in urllib.parse.parse_qs(u.query).items()}
            try:
                if u.path == "/":
                    return self._send(200, DASH_HTML, "text/html; charset=utf-8")
                if u.path == "/api/info":
                    return self._json({"empresas": empresas, "ultimo": _last_closed()})
                if u.path == "/api/generar":
                    emp = (q.get("empresa") or "DAE").upper()
                    if emp not in [e["code"] for e in empresas]:
                        return self._json({"error": f"empresa desconocida: {emp}"}, 400)
                    per = q.get("periodo") or None
                    if per and not re.fullmatch(r"20\d\d-\d\d", per):
                        return self._json({"error": "mes inválido (usa AAAA-MM)"}, 400)
                    with lock:  # un cálculo a la vez (abre el Excel)
                        return self._json(dash_generar(root, emp, per, meetings))
                if u.path == "/api/dynamics":
                    with lock:
                        D = dyn_build(root, "DAE", verbose=False, force=True)
                    return self._json({"fuente": (D or {}).get("fuente"), "actualizado": (D or {}).get("actualizado"),
                                       "ultimo_mov": (D or {}).get("ultimo_mov"), "aviso": (D or {}).get("aviso")})
                if u.path == "/api/preguntar":
                    with lock:
                        return self._json({"respuesta": answer_question(root, q.get("q", "")[:300], meetings)})
                return self._send(404, "no encontrado", "text/plain; charset=utf-8")
            except Exception as e:
                return self._json({"error": str(e)}, 500)

    srv = None
    for p in range(port, port + 10):
        try:
            srv = http.server.ThreadingHTTPServer(("127.0.0.1", p), Handler)  # solo tu computadora, nada sale a internet
            port = p
            break
        except OSError:
            continue
    if srv is None:
        sys.exit("No pude abrir el dashboard (puertos 8765-8774 ocupados).")
    url = f"http://127.0.0.1:{port}/"
    print(f"\nDashboard VEMO IA abierto en {url}\n(solo en tu computadora; deja esta ventana abierta y ciérrala para apagarlo)")
    if open_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        srv.server_close()


def main():
    fix_console()
    ensure_dirs()
    if len(sys.argv) == 1:
        try:
            menu()
        except (KeyboardInterrupt, EOFError):
            pass
        return
    import argparse
    ap = argparse.ArgumentParser(description="VEMO IA — boxes del MFR")
    ap.add_argument("--root", default=None)
    sub = ap.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("comentar")
    c.add_argument("--empresa", default="DAE")
    c.add_argument("--periodo", default=None)
    c.add_argument("--estilo", choices=["calidad", "decks"], default="calidad")
    c.add_argument("--aprender", choices=["auto", "si", "no"], default="auto")
    dy = sub.add_parser("dynamics")
    dy.add_argument("--configurar", action="store_true")
    dy.add_argument("--probar", action="store_true")
    dy.add_argument("--empresa", default="DAE")
    dsh = sub.add_parser("dashboard")
    dsh.add_argument("--puerto", type=int, default=DASH_PORT)
    dsh.add_argument("--sin-navegador", action="store_true")
    q = sub.add_parser("preguntar")
    q.add_argument("pregunta", nargs="*")
    a = sub.add_parser("aprender")
    a.add_argument("--empresa", default="DAE")
    a.add_argument("--examen", action="store_true", help="califica también en meses no vistos (más lento)")
    a.add_argument("--presupuesto-seg", type=int, default=None, help="(compatibilidad; ya no se usa)")
    args = ap.parse_args()
    root = resolve_root(args.root)
    if not root:
        sys.exit("No encuentro tu carpeta de FP&A: python vemo_ia.py  (sin argumentos) para configurarla.")
    if args.cmd == "comentar":
        emp = args.empresa.upper()
        if emp not in EJECUTA:
            sys.exit(f"Por ahora el motor MFR corre solo para {', '.join(EJECUTA)}.")
        run_comentar(root, emp, parse_periodo(args.periodo) if args.periodo else None, args.estilo, args.aprender)
    elif args.cmd == "dynamics":
        if args.configurar:
            dyn_api_configurar(args.empresa.upper())
        elif args.probar:
            dyn_api_probar(args.empresa.upper())
        else:
            dyn_build(root, args.empresa.upper(), verbose=True, force=True)
    elif args.cmd == "dashboard":
        run_dashboard(root, port=args.puerto, open_browser=not args.sin_navegador)
    elif args.cmd == "preguntar":
        if args.pregunta:
            print(_console(answer_question(root, " ".join(args.pregunta), get_meetings())))
        else:
            run_preguntar(root)
    else:
        if args.examen:
            os.environ["VEMO_MFR_CV"] = "1"
        learn_mfr(root, company=args.empresa.upper())


COMPANY_FILE_KEYWORDS = {
    "DAE": ["DAE"],
    "LTO": ["LTO"],
    "EVF": ["EVFLEETS"],
    "VCN": ["VCN"],
}
EXCLUDE_KEYWORDS = ["P&L_HUB", "PL_HUB", "BACKUP", "RESPALDO"]


if __name__ == "__main__":
    main()
