"""
WhatsApp chat parser + PPTX generator for Mondelez Milka implementations.
"""
import re, json, os, copy, zipfile, csv, difflib, unicodedata, urllib.request, io
from datetime import datetime
from pathlib import Path
from collections import defaultdict
from PIL import Image, ImageOps
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from lxml import etree


# ── Store database (Google Sheets) ───────────────────────────────────────────

# URL del CSV publicado – se puede sobreescribir con la variable de entorno STORE_DB_URL
SHEETS_CSV_URL = os.environ.get(
    'STORE_DB_URL',
    'https://docs.google.com/spreadsheets/d/e/'
    '2PACX-1vRX_MbNxlPJqTpAg89E51WOrp-oqNd6fAwjlN00ON6-SG1tGzQZNj7ZTs-0vgRAy53u0Dqjhi0I6Cyn'
    '/pub?output=csv'
)

_store_db_cache: list = None      # filas crudas
_store_db_index: list = None      # (nombre_norm, words_set, row) pre-calculado
_lookup_cache:   dict = {}        # query → resultado cacheado


def _norm(text: str) -> str:
    """Normaliza texto: minúsculas, sin tildes, sin puntuación."""
    text = text.lower().strip()
    text = unicodedata.normalize('NFD', text)
    text = ''.join(c for c in text if unicodedata.category(c) != 'Mn')
    text = re.sub(r'[^\w\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def load_store_db() -> list:
    """Carga y cachea la base de tiendas desde Google Sheets CSV."""
    global _store_db_cache, _store_db_index
    if _store_db_cache is not None:
        return _store_db_cache
    try:
        req = urllib.request.urlopen(SHEETS_CSV_URL, timeout=15)
        content = req.read().decode('utf-8-sig')
        rows = list(csv.reader(content.splitlines()))
        _store_db_cache = rows[1:]
        # Pre-normalizar una sola vez al cargar
        _store_db_index = [
            (_norm(row[5]), set(_norm(row[5]).split()), row)
            for row in _store_db_cache if len(row) >= 9
        ]
        print(f'[store_db] {len(_store_db_cache)} tiendas cargadas.', flush=True)
    except Exception as exc:
        print(f'[store_db] No se pudo cargar la base de tiendas: {exc}', flush=True)
        _store_db_cache = []
        _store_db_index = []
    return _store_db_cache


def lookup_store(query: str, chain: str = None):
    """
    Busca la tienda más parecida a *query* en el índice pre-normalizado.
    - chain: cadena detectada en WhatsApp (ej. 'SISA', 'JUMBO') para filtrar
             y evitar matches cruzados entre cadenas distintas.
    """
    cache_key = f"{chain}|{query}"
    if cache_key in _lookup_cache:
        return _lookup_cache[cache_key]

    load_store_db()
    if not _store_db_index:
        _lookup_cache[cache_key] = None
        return None

    q = _norm(query)
    q_words = set(q.split())

    # Prefijo de cadena para filtrar (primeras 4 letras normalizadas)
    chain_prefix = _norm(chain)[:4] if chain else None

    best_score, best_row = 0.0, None

    for nombre, nombre_words, row in _store_db_index:
        # Filtrar por cadena: el Nombre Sala debe empezar con el prefijo de la cadena
        if chain_prefix and not nombre.startswith(chain_prefix):
            continue
        # 1) Word-overlap rápido
        overlap = len(q_words & nombre_words) / max(len(q_words), 1)
        # 2) SequenceMatcher solo si overlap es prometedor
        if overlap > 0.3:
            ratio = difflib.SequenceMatcher(None, q, nombre).ratio()
        else:
            ratio = 0.0
        score = max(ratio, overlap * 0.88)
        if score > best_score:
            best_score = score
            best_row = row

    THRESHOLD = 0.55
    result = None
    if best_score >= THRESHOLD and best_row is not None:
        result = {
            'cadena':      best_row[3],
            'nombre_sala': best_row[5],
            'comuna':      best_row[7],
            'region':      best_row[8],
            'score':       round(best_score, 3),
        }
    _lookup_cache[cache_key] = result
    return result

CHAIN_ORDER = ['SISA', 'JUMBO', 'HIPER', 'SANTA ISABEL', 'TOTTUS', 'SMU', 'UNIMARC']

CHAIN_RULES = [
    ('SISA',         r'\bSisa\b|\bSISA\b',       'N'),
    ('JUMBO',        r'\bJUMBO\b|\bJumbo\b',      'J'),
    ('HIPER',        r'\bHIPER\b|\bHiper\b',       'H'),
    ('SANTA ISABEL', r'Santa Isabel|SANTA ISABEL', 'SI'),
    ('TOTTUS',       r'\bTOTTUS\b|\bTottus\b',    'T'),
    ('SMU',          r'\bSMU\b',                   'S'),
    ('UNIMARC',      r'\bUnimarc\b|\bUNIMARC\b',  'U'),
]

BAD_WORDS = [
    # Operativos / logística
    'implementación', 'implementacion', 'botadero', 'payloader', 'payloder',
    'material', 'ingreso', 'autorizar', 'abastece', 'abastecer', 'stock',
    'bodega', 'armar', 'solicita', 'solicitud', 'encargada', 'reponedor',
    'rechaza', 'pendiente', 'dejar', 'dejar en', 'nota en',
    # Comunicación / saludo
    'campaña', 'correo', 'problema', 'aparece', 'quieren', 'fecha de término',
    'agregué', 'buen día', 'estará', 'porfa', 'enviados', 'si no',
    'revisando', 'añadir', 'están', 'terminó', 'buen',
    # Señales de que es un reporte, no un identificador de tienda
    'imagen omitida', 'adjunto:', '@',
]

# Soporta formato 12h con AM/PM y formato 24h sin AM/PM
LINE_RE = re.compile(r'^\[(\d{2}-\d{2}-\d{2}), (\d{1,2}:\d{2}:\d{2})(?:\u202f([AP]M))?\] ([^:]+): (.*)$')
# Adjuntos formato antiguo: <attached: file.jpg> / <adjunto: file.jpg>
ATTACH_RE = re.compile(
    r'<(?:attached|adjunto):\s*([^\s>][^>]*\.(?:jpg|jpeg|png|webp))\s*>',
    re.IGNORECASE
)
# Adjuntos formato nuevo: file.jpg (file attached) / file.jpg (archivo adjunto)
ATTACH_RE_PAREN = re.compile(
    r'([\w][\w\-\.]*\.(?:jpg|jpeg|png|webp))\s*\((?:file attached|archivo adjunto)\)',
    re.IGNORECASE
)


def find_photos(text: str) -> list:
    """Extrae nombres de archivo de fotos de un fragmento de texto WhatsApp.
    Soporta formato antiguo (<adjunto: ...>) y formato nuevo (... (archivo adjunto)).
    """
    return ATTACH_RE.findall(text) + ATTACH_RE_PAREN.findall(text)


# ── Chat parsing ──────────────────────────────────────────────────────────────

def parse_messages(chat_text: str) -> list:
    content = chat_text.replace('\u200e', '').replace('\u200f', '').replace('\r', '')
    messages, current = [], None
    for line in content.split('\n'):
        m = LINE_RE.match(line)
        if m:
            if current:
                messages.append(current)
            d, t, ap, sender, txt = m.groups()
            if ap:  # formato 12h con AM/PM
                dt = datetime.strptime(f'{d} {t} {ap}', '%d-%m-%y %I:%M:%S %p')
            else:   # formato 24h sin AM/PM
                dt = datetime.strptime(f'{d} {t}', '%d-%m-%y %H:%M:%S')
            current = {'dt': dt, 'sender': sender.strip(), 'text': txt,
                       'photos': find_photos(txt)}
        else:
            if current:
                current['text'] += '\n' + line
                current['photos'] += find_photos(line)
    if current:
        messages.append(current)
    return messages


def detect_chain(text):
    first = text.strip().split('\n')[0]
    for chain, pat, prefix in CHAIN_RULES:
        if re.search(pat, first, re.IGNORECASE):
            return chain, prefix
    return None, None


def is_store_message(text):
    chain, _ = detect_chain(text)
    if not chain:
        return False
    first = text.strip().split('\n')[0].lower()
    # Rechazar si contiene palabras de reporte/logística
    if any(b in first for b in BAD_WORDS):
        return False
    # Rechazar patrón "CADENA - [texto libre]" → es un reporte, no identificador
    # Ejemplo: "SISA - Implementación botadero..." / "JUMBO - En Independencia..."
    chain_stripped = re.sub(
        r'\b(?:sisa|jumbo|hiper|santa isabel|tottus|smu|unimarc)\b', '', first, flags=re.IGNORECASE
    ).strip()
    if chain_stripped.startswith('- ') or chain_stripped.startswith('-\t'):
        return False
    return True


def parse_store_line(text, chain, prefix):
    first = re.sub(r'<(?:attached|adjunto):[^>]+>', '', text.strip().split('\n')[0]).strip()
    parts = first.split('\t')
    if len(parts) >= 3:
        raw_code = re.sub(r'[°º]', '', parts[0].strip())
        addr = parts[2].strip()
        city = parts[3].strip() if len(parts) > 3 else ''
        if re.match(r'^[A-Za-z]\d+$', raw_code):
            code = raw_code.upper()
        elif re.match(r'^\d+$', raw_code):
            code = prefix + raw_code
        else:
            code = raw_code.upper()
        return code, addr, city
    # Free-form
    stripped = first
    for _, pat, _ in CHAIN_RULES:
        stripped = re.sub(pat, '', stripped, flags=re.IGNORECASE).strip(' ,\t')
    m = re.match(r'^(\d+)\s+(.+)$', stripped)
    if m:
        return prefix + m.group(1), m.group(2).strip(), ''
    return None, stripped.strip(), ''


def parse_status(text):
    pl, bt, notes = None, 0, []
    for line in text.split('\n'):
        l, lu = line.strip(), line.strip().upper()
        if '\u2705' in l:
            if 'PAYLOAD' in lu:
                pl = 'Implementado'
            if 'BOTADERO' in lu:
                mm = re.search(r'(\d+)\s*BOTADERO', lu)
                bt = int(mm.group(1)) if mm else 1
        if '\u274c' in l:
            if 'PAYLOAD' in lu:
                pl = 'No implementado'
            note = re.sub(r'[\u274c\u2705]', '', l).strip()
            if note:
                notes.append(note)
        if 'NO SE PUDO IMPLEMENTAR' in lu and 'PAYLOAD' in lu:
            pl = 'No implementado'
    return pl, bt, ' | '.join(notes)


def extract_stores(messages: list, start_date: datetime, end_date: datetime) -> list:
    recent = [m for m in messages if start_date <= m['dt'] < end_date]
    raw = []
    for i, msg in enumerate(recent):
        if not is_store_message(msg['text']):
            continue
        chain, prefix = detect_chain(msg['text'])
        code, address, city = parse_store_line(msg['text'], chain, prefix)
        photos = []
        for j in range(max(0, i - 20), min(len(recent), i + 10)):
            other = recent[j]
            if other['sender'] != msg['sender']:
                continue
            diff = (other['dt'] - msg['dt']).total_seconds()
            if -180 <= diff <= 180 and other['photos']:
                photos.extend(other['photos'])
        seen = set()
        photos = [p for p in photos if not (p in seen or seen.add(p))]
        pl, bt, notes = parse_status(msg['text'])
        print(f'[extract] tienda={code or address[:30]!r} fotos={len(photos)}', flush=True)

        # Buscar tienda en base de datos formal (filtrando por cadena)
        query = f"{chain} {address}" if address else chain
        db_match = lookup_store(query, chain=chain)

        raw.append({
            'chain': chain, 'code': code, 'address': address, 'city': city,
            'sender': msg['sender'], 'date': msg['dt'].strftime('%d/%m/%Y'),
            'datetime': msg['dt'].isoformat(),
            'photos': photos, 'payloader': pl, 'botaderos': bt, 'notes': notes,
            # Datos formales desde la planilla (None si no hubo match)
            'db_cadena':      db_match['cadena']      if db_match else None,
            'db_nombre_sala': db_match['nombre_sala']  if db_match else None,
            'db_comuna':      db_match['comuna']       if db_match else None,
            'db_region':      db_match['region']       if db_match else None,
        })
    # Deduplicate
    deduped = {}
    for s in raw:
        key = f"{s['chain']}_{s['code'] or s['address'][:20]}"
        if key in deduped:
            ex = deduped[key]
            merged = ex['photos'] + s['photos']
            seen = set()
            ex['photos'] = [p for p in merged if not (p in seen or seen.add(p))]
            if s['payloader']:
                ex['payloader'] = s['payloader']
            if s['botaderos']:
                ex['botaderos'] = s['botaderos']
        else:
            deduped[key] = s
    stores = list(deduped.values())
    stores.sort(key=lambda s: (
        CHAIN_ORDER.index(s['chain']) if s['chain'] in CHAIN_ORDER else 99,
        s['datetime'],
        s['code'] or ''
    ))
    return stores


# ── PPTX generation ───────────────────────────────────────────────────────────

def open_corrected(img_path: str, max_px: int = 1200):
    """
    Corrige orientación EXIF y reduce resolución solo si es necesario.
    - Si la foto ya es pequeña y no está rotada → devuelve None (usa archivo original, rápido).
    - Si necesita ajuste → devuelve BytesIO con imagen corregida.
    """
    try:
        img = Image.open(img_path)   # lazy: no decodifica píxeles todavía

        # Leer orientación EXIF sin decodificar la imagen completa
        try:
            orientation = (img.getexif() or {}).get(274, 1)
        except Exception:
            orientation = 1

        needs_rotate = orientation not in (1, 0, None)
        needs_resize = img.width > max_px or img.height > max_px

        if not needs_rotate and not needs_resize:
            return None   # sin cambios → usar archivo original directamente

        img.load()   # decodificar solo si realmente hay que modificar
        img = ImageOps.exif_transpose(img)
        if img.mode not in ('RGB', 'L'):
            img = img.convert('RGB')
        if img.width > max_px or img.height > max_px:
            img.thumbnail((max_px, max_px), Image.LANCZOS)

        buf = io.BytesIO()
        img.save(buf, format='JPEG', quality=82)
        buf.seek(0)
        return buf
    except Exception:
        return None


def build_photo_index(photos_dir: str) -> dict:
    """
    Construye un índice {nombre_archivo → path_completo} buscando recursivamente
    en photos_dir. Necesario cuando el ZIP tiene subcarpetas (ej. Media/).
    """
    index = {}
    for root, _, files in os.walk(photos_dir):
        for f in files:
            if f.lower().endswith(('.jpg', '.jpeg', '.png', '.webp')):
                # Guardamos solo si no está ya (prioridad a rutas más cortas / raíz)
                if f not in index:
                    index[f] = os.path.join(root, f)
    return index


def select_photos(photos, n, photos_dir, _photo_index=None):
    """
    Selecciona hasta n fotos de la lista y devuelve sus paths completos.
    Usa _photo_index si ya fue construido; si no, lo construye internamente.
    """
    if n == 0:
        return []
    idx = _photo_index if _photo_index is not None else build_photo_index(photos_dir)
    avail = [idx[p] for p in photos if p in idx]
    if not avail:
        return []
    if len(avail) <= n:
        return avail
    step = len(avail) / n
    return [avail[int(i * step)] for i in range(n)]


def group_runs_by_br(para):
    groups, run_idx = [[]], 0
    for child in para._p:
        tag = etree.QName(child).localname
        if tag == 'r':
            groups[-1].append(run_idx)
            run_idx += 1
        elif tag == 'br':
            groups.append([])
    return groups


def set_line(para, idxs, text):
    runs = para.runs
    if not idxs:
        return
    if idxs[0] < len(runs):
        runs[idxs[0]].text = text
    for i in idxs[1:]:
        if i < len(runs):
            runs[i].text = ''


def update_caption(shape, slot, fecha, pl_status, bt_status):
    element = 'Payloader Easter' if slot == 0 else 'Botadero Easter'
    status = pl_status if slot == 0 else bt_status
    tf = shape.text_frame
    if not tf.paragraphs:
        return
    para = tf.paragraphs[0]
    groups = group_runs_by_br(para)
    n = len(groups)
    if n >= 1:
        set_line(para, groups[0], f'FOTO IMPLEMENTACIÓN {slot + 1}')
    if n >= 2:
        set_line(para, groups[1], f'FECHA: {fecha}')
    if n >= 3:
        g, runs = groups[2], para.runs
        if len(g) >= 3:
            runs[g[0]].text = 'ELEMENTO: '
            runs[g[1]].text = 'Payloader' if slot == 0 else 'Botadero'
            runs[g[2]].text = ' Easter'
            for x in g[3:]:
                runs[x].text = ''
        elif g:
            runs[g[0]].text = f'ELEMENTO: {element}'
            for x in g[1:]:
                runs[x].text = ''
    if n >= 4:
        set_line(para, groups[3], f'STATUS: {status}')


def update_store_slide(slide, store, photos_dir, photo_index=None):
    code = store.get('code', '')
    address = store.get('address', '')
    city = store.get('city', '')
    chain = store.get('chain', '')
    fecha = store.get('date', '--/--/----')
    photos = store.get('photos', [])
    pl_stat = store.get('payloader') or 'Implementado'
    bt_stat = 'Implementado'

    chain_label = chain if chain else 'SISA'

    # Usar datos formales de la planilla si están disponibles
    db_nombre = store.get('db_nombre_sala')
    db_comuna = store.get('db_comuna')
    db_region = store.get('db_region')
    db_cadena = store.get('db_cadena')

    if db_nombre:
        # Formato: NOMBRE SALA — Comuna, Región
        partes = [db_nombre]
        if db_comuna:
            partes.append(db_comuna)
        if db_region:
            partes.append(db_region)
        header_text = ' — '.join(partes[:1]) + (f' — {db_comuna}' if db_comuna else '')
        if db_region and db_region != db_comuna:
            header_text += f', {db_region}'
    else:
        # Fallback: datos parseados desde WhatsApp
        header_text = (f'{code} {chain_label} - {address}' if code else f'{chain_label} - {address}')
        if city:
            header_text += f', {city}'

    text_shapes = [s for s in slide.shapes if s.has_text_frame]

    # Image containers: PICTURE (13) or image PLACEHOLDER (14 without text)
    pic_shapes = sorted(
        [s for s in slide.shapes if s.shape_type == 13 or
         (s.shape_type == 14 and not s.has_text_frame)],
        key=lambda s: s.left
    )

    # Header: widest text shape NOT containing caption keywords
    CAPTION_KEYS = ('FOTO', 'FECHA', 'ELEMENTO', 'STATUS')
    for sh in sorted(text_shapes, key=lambda s: -s.width):
        t = sh.text_frame.text
        if not any(k in t for k in CAPTION_KEYS):
            for para in sh.text_frame.paragraphs:
                if para.runs:
                    para.runs[0].text = header_text
                    for r in para.runs[1:]:
                        r.text = ''
                    break
            break

    captions = sorted([s for s in text_shapes
                       if any(k in s.text_frame.text for k in ('FOTO', 'FECHA', 'ELEMENTO'))],
                      key=lambda s: s.left)
    for slot, cap in enumerate(captions):
        update_caption(cap, slot, fecha, pl_stat, bt_stat)

    sel = select_photos(photos, len(pic_shapes), photos_dir, _photo_index=photo_index)
    sorted_pics = sorted(pic_shapes, key=lambda s: s.left)
    for i, pic_sh in enumerate(sorted_pics):
        left, top, w, h = pic_sh.left, pic_sh.top, pic_sh.width, pic_sh.height
        slide.shapes._spTree.remove(pic_sh._element)
        if i < len(sel):
            img_path = sel[i]  # ya es path completo
            try:
                img_src = open_corrected(img_path) or img_path
                slide.shapes.add_picture(img_src, left, top, w, h)
            except Exception:
                pass
        else:
            if i < len(captions):
                try:
                    slide.shapes._spTree.remove(captions[i]._element)
                except Exception:
                    pass


def add_slide_copy(prs, src_idx):
    src = prs.slides[src_idx]
    new = prs.slides.add_slide(src.slide_layout)
    for shape in list(new.shapes):
        new.shapes._spTree.remove(shape._element)
    for shape in src.shapes:
        new.shapes._spTree.append(copy.deepcopy(shape._element))
    return new


def make_chain_divider(prs, chain_name, title_slide_idx=0):
    new = add_slide_copy(prs, title_slide_idx)
    for sh in new.shapes:
        if sh.has_text_frame:
            t = sh.text_frame.text
            if 'IMPLEMENTACIÓN' in t or 'MILKA' in t:
                paras = sh.text_frame.paragraphs
                if paras[0].runs:
                    paras[0].runs[0].text = chain_name
                    for r in paras[0].runs[1:]:
                        r.text = ''
                for para in paras[1:]:
                    for r in para.runs:
                        r.text = ''
                break
    return new


def generate_pptx(stores: list, photos_dir: str, template_path: str, output_path: str) -> dict:
    """Generate the combined PPTX. Returns summary dict."""
    prs = Presentation(template_path)

    # Keep only title (0) and one store slide as template (1)
    keep = [0, 1]
    remove = sorted(set(range(len(prs.slides))) - set(keep), reverse=True)
    for idx in remove:
        rId = prs.slides._sldIdLst[idx].get(
            '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[idx]

    # Construir índice de fotos una sola vez (búsqueda recursiva en subdirectorios)
    photo_index = build_photo_index(photos_dir)
    print(f'[pptx] {len(photo_index)} fotos indexadas en {photos_dir}', flush=True)

    by_chain = defaultdict(list)
    for s in stores:
        by_chain[s['chain']].append(s)

    summary = {}
    for chain in CHAIN_ORDER:
        chain_stores = by_chain.get(chain, [])
        if not chain_stores:
            continue
        make_chain_divider(prs, chain, title_slide_idx=0)
        for store in chain_stores:
            new_slide = add_slide_copy(prs, 1)
            update_store_slide(new_slide, store, photos_dir, photo_index=photo_index)
        summary[chain] = len(chain_stores)

    # Remove the store template slide
    rId = prs.slides._sldIdLst[1].get(
        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[1]

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    prs.save(output_path)
    return summary


def process_zip(zip_path: str, photos_dir: str, start_date: datetime,
                end_date: datetime, template_path: str, output_path: str) -> dict:
    """Full pipeline: unzip → parse → generate PPTX. Returns result dict."""
    # Extract ZIP
    with zipfile.ZipFile(zip_path, 'r') as zf:
        zf.extractall(photos_dir)

    # Find _chat.txt
    chat_file = Path(photos_dir) / '_chat.txt'
    if not chat_file.exists():
        raise FileNotFoundError('No se encontró _chat.txt en el ZIP')

    chat_text = chat_file.read_text(encoding='utf-8', errors='replace')
    messages = parse_messages(chat_text)
    stores = extract_stores(messages, start_date, end_date)

    if not stores:
        raise ValueError('No se encontraron tiendas en el rango de fechas seleccionado')

    summary = generate_pptx(stores, photos_dir, template_path, output_path)
    total_slides = 1 + sum(len(v) + 1 for v in defaultdict(list,
                           {c: [s for s in stores if s['chain'] == c]
                            for c in CHAIN_ORDER}).values() if v)
    return {
        'stores': stores,
        'summary': summary,
        'total_slides': total_slides,
        'output_path': output_path,
    }
