"""
WhatsApp chat parser + PPTX generator for Mondelez Milka implementations.
"""
import re, json, os, copy, zipfile
from datetime import datetime
from pathlib import Path
from collections import defaultdict
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from lxml import etree

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

BAD_WORDS = ['campaña', 'correo', 'problema', 'aparece', 'quieren', 'fecha de término',
             'agregué', 'buen día', 'estará', 'porfa', 'enviados', 'si no',
             'revisando', 'añadir', 'están', 'terminó', 'implementación?', 'buen']

LINE_RE   = re.compile(r'^\[(\d{2}-\d{2}-\d{2}), (\d{1,2}:\d{2}:\d{2})\u202f([AP]M)\] ([^:]+): (.*)$')
ATTACH_RE = re.compile(r'<attached: (\d+-PHOTO-[\d-]+\.jpg)>')


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
            dt = datetime.strptime(f'{d} {t} {ap}', '%d-%m-%y %I:%M:%S %p')
            current = {'dt': dt, 'sender': sender.strip(), 'text': txt,
                       'photos': ATTACH_RE.findall(txt)}
        else:
            if current:
                current['text'] += '\n' + line
                current['photos'] += ATTACH_RE.findall(line)
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
    return not any(b in first for b in BAD_WORDS)


def parse_store_line(text, chain, prefix):
    first = re.sub(r'<attached:[^>]+>', '', text.strip().split('\n')[0]).strip()
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
        raw.append({
            'chain': chain, 'code': code, 'address': address, 'city': city,
            'sender': msg['sender'], 'date': msg['dt'].strftime('%d/%m/%Y'),
            'datetime': msg['dt'].isoformat(),
            'photos': photos, 'payloader': pl, 'botaderos': bt, 'notes': notes,
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

def select_photos(photos, n, photos_dir):
    if n == 0:
        return []
    avail = [p for p in photos if Path(photos_dir, p).exists()]
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


def update_store_slide(slide, store, photos_dir):
    code = store.get('code', '')
    address = store.get('address', '')
    city = store.get('city', '')
    chain = store.get('chain', '')
    fecha = store.get('date', '--/--/----')
    photos = store.get('photos', [])
    pl_stat = store.get('payloader') or 'Implementado'
    bt_stat = 'Implementado'

    chain_label = chain if chain else 'SISA'
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

    sel = select_photos(photos, len(pic_shapes), photos_dir)
    sorted_pics = sorted(pic_shapes, key=lambda s: s.left)
    for i, pic_sh in enumerate(sorted_pics):
        left, top, w, h = pic_sh.left, pic_sh.top, pic_sh.width, pic_sh.height
        slide.shapes._spTree.remove(pic_sh._element)
        if i < len(sel):
            img = str(Path(photos_dir, sel[i]))
            try:
                slide.shapes.add_picture(img, left, top, w, h)
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
            update_store_slide(new_slide, store, photos_dir)
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
