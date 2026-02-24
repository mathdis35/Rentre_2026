import os, re, shutil, datetime, uuid, json, copy, calendar
from pathlib import Path
from collections import defaultdict
from flask import Flask, render_template, request, jsonify, send_file, after_this_request

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    import xlrd
except ImportError:
    pass

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
UPLOAD_FOLDER = '/tmp/plannipro'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

MOIS_ABBR = {
    'août':8,'aout':8,'sept':9,'septembre':9,'oct':10,'octobre':10,
    'nov':11,'novembre':11,'déc':12,'dec':12,'décembre':12,
    'janv':1,'jan':1,'janvier':1,'févr':2,'fev':2,'février':2,
    'mars':3,'avr':4,'avril':4,'mai':5,'juin':6,'juil':7,'juillet':7
}
MOIS_FR = ['','Janvier','Février','Mars','Avril','Mai','Juin',
           'Juillet','Août','Septembre','Octobre','Novembre','Décembre']

FERIES = {
    datetime.date(2026,11,1), datetime.date(2026,11,11), datetime.date(2026,12,25),
    datetime.date(2027,1,1),  datetime.date(2027,4,5),   datetime.date(2027,5,1),
    datetime.date(2027,5,8),  datetime.date(2027,5,13),  datetime.date(2027,5,24),
    datetime.date(2027,7,14), datetime.date(2027,8,15),
}

DEFAULT_COLORS = ['FFEE7E32','FFFFCC99','FFFFD243','FFDDFFDD',
                  'FFB8D4F0','FFFFC0CB','FFD4B8F0','FFB8F0D4']

# ── Copie de styles compatible Python 3.14 ────
def copy_cell_style(src, dst):
    """Copie les styles sans copy.copy() — incompatible avec Python 3.14 + openpyxl"""
    try:
        f = src.font
        dst.font = Font(
            name=f.name, size=f.size, bold=f.bold, italic=f.italic,
            underline=f.underline, color=f.color, strike=f.strike,
            vertAlign=f.vertAlign
        )
    except: pass
    try:
        fi = src.fill
        if fi and fi.fill_type == 'solid':
            fg = fi.fgColor.rgb if fi.fgColor and fi.fgColor.type == 'rgb' else 'FFFFFFFF'
            dst.fill = PatternFill(start_color=fg, end_color=fg, fill_type='solid')
        elif fi and fi.fill_type and fi.fill_type != 'none':
            dst.fill = PatternFill(fill_type=fi.fill_type)
    except: pass
    try:
        b = src.border
        def copy_side(s):
            if s is None: return Side()
            return Side(style=s.style, color=s.color)
        dst.border = Border(
            left=copy_side(b.left), right=copy_side(b.right),
            top=copy_side(b.top), bottom=copy_side(b.bottom)
        )
    except: pass
    try:
        a = src.alignment
        dst.alignment = Alignment(
            horizontal=a.horizontal, vertical=a.vertical,
            wrap_text=a.wrap_text, shrink_to_fit=a.shrink_to_fit,
            indent=a.indent, text_rotation=a.text_rotation
        )
    except: pass
    try:
        dst.number_format = src.number_format
    except: pass

# ── Helpers ───────────────────────────────────
def colors_match(rgb):
    if not rgb: return False
    return abs(rgb[0]-166)<20 and abs(rgb[1]-202)<20 and abs(rgb[2]-240)<20

def cell_to_str(v):
    if v is None: return ''
    return str(v).strip()

def is_available(v):
    if v is None: return False
    return str(v).strip().upper() in ['X','4','✓','OUI','O']

def find_month_num(text):
    t = text.strip().lower()
    for m, mn in MOIS_ABBR.items():
        if m in t:
            yr = re.search(r'(20\d\d)', t)
            return mn, int(yr.group(1)) if yr else None
    return None, None

# ── Parser planning classe XLS ────────────────
def parse_planning_xls(filepath):
    try:
        wb = xlrd.open_workbook(filepath, formatting_info=True)
    except Exception as e:
        return {'nom': Path(filepath).stem, 'jours': []}
    ws = wb.sheet_by_index(0)
    nom = Path(filepath).stem
    for c in range(ws.ncols):
        v = ws.cell_value(0, c)
        if isinstance(v, str) and len(v.strip()) > 3:
            nom = v.strip(); break

    jours = []
    blocks = []
    for c in range(ws.ncols):
        v = ws.cell_value(4, c)
        if isinstance(v, str):
            mn, yr = find_month_num(v)
            if mn and yr:
                blocks.append({'cj': c+1, 'cd': c+2, 'y': yr, 'm': mn})

    for b in blocks:
        for r in range(5, ws.nrows):
            if b['cj'] >= ws.ncols or b['cd'] >= ws.ncols: continue
            try:
                xf = wb.xf_list[ws.cell_xf_index(r, b['cj'])]
                rgb = wb.colour_map.get(xf.background.pattern_colour_index)
                if not colors_match(rgb): continue
            except: continue
            dv = ws.cell_value(r, b['cd'])
            dn = int(dv) if isinstance(dv, float) and dv > 0 else None
            if isinstance(dv, str) and dv.strip().isdigit(): dn = int(dv.strip())
            if not dn or not (1 <= dn <= 31): continue
            try:
                d = datetime.date(b['y'], b['m'], dn)
                if d.weekday() < 5: jours.append(d.strftime('%Y-%m-%d'))
            except: pass
    return {'nom': nom, 'jours': sorted(set(jours))}

# ── Parser planning classe XLSX ───────────────
def parse_planning_xlsx(filepath):
    try:
        wb = load_workbook(filepath, data_only=True)
    except: return {'nom': Path(filepath).stem, 'jours': []}
    ws = wb.active
    nom = Path(filepath).stem
    for c in range(1, ws.max_column+1):
        v = ws.cell(row=1, column=c).value
        if v and isinstance(v, str) and len(v.strip()) > 3:
            nom = v.strip(); break

    jours = []
    blocks = []
    for c in range(1, ws.max_column+1):
        v = ws.cell(row=5, column=c).value
        if isinstance(v, str):
            mn, yr = find_month_num(v)
            if mn and yr:
                blocks.append({'cj': c+1, 'cd': c+2, 'y': yr, 'm': mn})

    for b in blocks:
        for r in range(6, ws.max_row+1):
            cell = ws.cell(row=r, column=b['cj'])
            bg = cell.fill.fgColor.rgb if cell.fill and cell.fill.fgColor and cell.fill.fgColor.type == 'rgb' else None
            if bg not in ('FFA6CAF0', 'A6CAF0'): continue
            dv = ws.cell(row=r, column=b['cd']).value
            dn = None
            if isinstance(dv, datetime.datetime): dn = dv.day
            elif isinstance(dv, (int, float)): dn = int(dv)
            elif isinstance(dv, str) and dv.strip().isdigit(): dn = int(dv.strip())
            if not dn or not (1 <= dn <= 31): continue
            try:
                d = datetime.date(b['y'], b['m'], dn)
                if d.weekday() < 5: jours.append(d.strftime('%Y-%m-%d'))
            except: pass
    return {'nom': nom, 'jours': sorted(set(jours))}

def parse_planning_classe(fp):
    return parse_planning_xls(fp) if Path(fp).suffix.lower() == '.xls' else parse_planning_xlsx(fp)

# ── Parser disponibilités ─────────────────────
def parse_disponibilite(filepath):
    nom = Path(filepath).stem
    dispo = {}
    try:
        wb = load_workbook(filepath, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
    except:
        return {'nom': nom, 'dispo': {}}

    if len(rows) < 2: return {'nom': nom, 'dispo': {}}

    header = rows[0]
    month_cols = []
    for ci, val in enumerate(header):
        if isinstance(val, str):
            mn, yr = find_month_num(val)
            if mn and yr:
                month_cols.append({'ci': ci, 'm': mn, 'y': yr})

    ign = {'sam', 'dim', 'férié', 'ferie', 'férie', 'fériés', 'nan', 'none', ''}

    for mc in month_cols:
        ci, mn, yr = mc['ci'], mc['m'], mc['y']
        col_matin = ci
        col_pm    = ci + 1
        for ri in range(2, len(rows)):
            row = rows[ri]
            day_abbr = cell_to_str(row[0]).lower() if len(row) > 0 else ''
            if day_abbr in ign: continue
            day_num_raw = row[1] if len(row) > 1 else None
            dn = None
            if isinstance(day_num_raw, (int, float)) and day_num_raw == day_num_raw:
                dn = int(day_num_raw)
            elif isinstance(day_num_raw, str) and day_num_raw.strip().isdigit():
                dn = int(day_num_raw.strip())
            if not dn or not (1 <= dn <= 31): continue
            try:
                d = datetime.date(yr, mn, dn)
                if d.weekday() >= 5: continue
                ds = d.strftime('%Y-%m-%d')
                mv = row[col_matin] if col_matin < len(row) else None
                pv = row[col_pm]    if col_pm    < len(row) else None
                if cell_to_str(mv).lower() in ign: continue
                matin = is_available(mv)
                pm    = is_available(pv)
                if matin or pm:
                    dispo[ds] = {'matin': matin, 'pm': pm}
            except: pass
    return {'nom': nom, 'dispo': dispo}

# ── Parser tableau formateurs ──────────────────
def parse_tableau_formateurs(filepath):
    try:
        wb = load_workbook(filepath, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
    except:
        return {}

    aff = {}
    if not rows: return aff
    header = [cell_to_str(c).lower() for c in rows[0]]

    col_nom = col_mat = col_pm = col_classe = None
    for i, h in enumerate(header):
        if 'nom' in h or 'formateur' in h: col_nom = i
        if 'matin' in h: col_mat = i
        if 'après' in h or 'apres' in h or 'pm' in h: col_pm = i
        if 'class' in h or 'groupe' in h: col_classe = i

    if col_nom is None: return aff

    for row in rows[1:]:
        nom = cell_to_str(row[col_nom]) if col_nom < len(row) else ''
        if not nom: continue
        matin   = is_available(row[col_mat])    if col_mat    is not None and col_mat    < len(row) else False
        pm      = is_available(row[col_pm])     if col_pm     is not None and col_pm     < len(row) else False
        classes = cell_to_str(row[col_classe])  if col_classe is not None and col_classe < len(row) else ''
        aff[nom] = {'matin': matin, 'pm': pm, 'classes': classes}
    return aff

# ── Assignation ───────────────────────────────
def assigner(planning_classes, dispos, aff):
    assignment = defaultdict(lambda: defaultdict(dict))
    stats = {'assigned': 0, 'warn': 0}
    heures = defaultdict(lambda: defaultdict(float))
    charge = defaultdict(float)

    for pc in planning_classes:
        for ds in pc['jours']:
            for demi in ['matin', 'pm']:
                candidats = []
                for d in dispos:
                    nom = d['nom']
                    if ds not in d['dispo']: continue
                    if not d['dispo'][ds].get(demi): continue
                    candidats.append(nom)
                if not candidats:
                    stats['warn'] += 1
                    continue
                candidats.sort(key=lambda n: charge[n])
                choisi = candidats[0]
                assignment[ds][demi][pc['nom']] = choisi
                charge[choisi] += 0.5
                heures[choisi][ds[:7]] = heures[choisi].get(ds[:7], 0) + 0.5
                stats['assigned'] += 1

    return assignment, stats, heures

# ── Écriture planning ─────────────────────────
def ecrire_planning(template_path, assignment, mois_sel, output_path):
    wb = load_workbook(template_path)
    for ws in wb.worksheets:
        titre = ws.title.lower()
        m_match = None
        for mn, mn_fr in enumerate(MOIS_FR):
            if mn_fr.lower() in titre:
                m_match = mn; break
        if m_match is None: continue
        yr_match = re.search(r'20\d\d', ws.title)
        if not yr_match: continue
        yr = int(yr_match.group())
        mk = f"{m_match}/{yr}"
        if mk not in mois_sel: continue
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str): continue
                ds_match = re.search(r'(\d{4}-\d{2}-\d{2})', str(cell.value))
                if not ds_match: continue
                ds = ds_match.group(1)
                if ds not in assignment: continue
                texts = []
                for demi in ['matin', 'pm']:
                    if demi in assignment[ds]:
                        for classe, formateur in assignment[ds][demi].items():
                            texts.append(f"{demi[0].upper()} – {formateur}")
                if texts:
                    cell.value = '\n'.join(texts)
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    break
    wb.save(output_path)

# ── Génération template colorié ───────────────
def detect_structure(ws):
    kw = ['BTS','BAC','EC ','CGC','NDRC','GPME','RDC','RH','Master']
    cc = {}; colors = {}
    for ri in range(1, 6):
        for ci in range(1, ws.max_column+1):
            cell = ws.cell(row=ri, column=ci); v = cell.value
            if isinstance(v, str) and any(k in v for k in kw):
                nom = v.strip(); cc[nom] = ci
                if cell.fill and cell.fill.fgColor and cell.fill.fgColor.type == 'rgb':
                    c = cell.fill.fgColor.rgb
                    if c not in ('00000000', 'FFFFFFFF'): colors[nom] = c
    for nom, ci in cc.items():
        if nom in colors: continue
        for ri in range(5, min(25, ws.max_row+1)):
            cell = ws.cell(row=ri, column=ci)
            if cell.fill and cell.fill.fgColor and cell.fill.fgColor.type == 'rgb':
                c = cell.fill.fgColor.rgb
                if c not in ('00000000', 'FFFFFFFF', None):
                    colors[nom] = c; break
    jf_set = {'lundi','mardi','mercredi','jeudi','vendredi','lun','mar','mer','jeu','ven'}
    fdr = 6; dlc = 2
    for ri in range(3, 20):
        for ci in range(1, 5):
            v = ws.cell(row=ri, column=ci).value
            if isinstance(v, str) and v.strip().lower() in jf_set:
                fdr = ri; dlc = ci; break
        else: continue
        break
    return {'class_cols': cc, 'class_colors': colors, 'first_data_row': fdr, 'day_label_col': dlc}

def snapshot_template(ws_src):
    """Lit toutes les données du template en mémoire (dict) — évite de recharger le fichier 12x"""
    snap = {'cells': {}, 'col_widths': {}, 'row_heights': {}, 'merges': []}
    for row in ws_src.iter_rows():
        for cell in row:
            entry = {'value': cell.value, 'style': None}
            if cell.has_style:
                try:
                    fi = cell.fill
                    fill = None
                    if fi and fi.fill_type == 'solid':
                        fg = fi.fgColor.rgb if fi.fgColor and fi.fgColor.type == 'rgb' else 'FFFFFFFF'
                        fill = ('solid', fg)
                    f = cell.font
                    font = {
                        'name': f.name, 'size': f.size, 'bold': f.bold,
                        'italic': f.italic, 'underline': f.underline,
                        'strike': f.strike, 'color': f.color, 'vertAlign': f.vertAlign
                    }
                    b = cell.border
                    def ss(s):
                        if s is None: return (None, None)
                        return (s.style, s.color)
                    border = (ss(b.left), ss(b.right), ss(b.top), ss(b.bottom))
                    a = cell.alignment
                    align = {
                        'horizontal': a.horizontal, 'vertical': a.vertical,
                        'wrap_text': a.wrap_text, 'shrink_to_fit': a.shrink_to_fit,
                        'indent': a.indent, 'text_rotation': a.text_rotation
                    }
                    entry['style'] = {'font': font, 'fill': fill, 'border': border,
                                      'align': align, 'number_format': cell.number_format}
                except: pass
            snap['cells'][(cell.row, cell.column)] = entry
    for cl, cd in ws_src.column_dimensions.items():
        snap['col_widths'][cl] = cd.width
    for ri, rd in ws_src.row_dimensions.items():
        snap['row_heights'][ri] = rd.height
    snap['merges'] = [str(mg) for mg in ws_src.merged_cells.ranges]
    snap['max_row'] = ws_src.max_row
    snap['max_col'] = ws_src.max_column
    return snap

def apply_snapshot(ws, snap):
    """Applique le snapshot sur une feuille vierge"""
    for (r, c), entry in snap['cells'].items():
        nc = ws.cell(row=r, column=c)
        nc.value = entry['value']
        s = entry.get('style')
        if not s: continue
        try:
            f = s['font']
            nc.font = Font(name=f['name'], size=f['size'], bold=f['bold'],
                           italic=f['italic'], underline=f['underline'],
                           strike=f['strike'], color=f['color'], vertAlign=f['vertAlign'])
        except: pass
        try:
            fill = s['fill']
            if fill and fill[0] == 'solid':
                nc.fill = PatternFill(start_color=fill[1], end_color=fill[1], fill_type='solid')
        except: pass
        try:
            bl, br, bt, bb = s['border']
            def ms(t):
                return Side(style=t[0], color=t[1]) if t[0] else Side()
            nc.border = Border(left=ms(bl), right=ms(br), top=ms(bt), bottom=ms(bb))
        except: pass
        try:
            a = s['align']
            nc.alignment = Alignment(horizontal=a['horizontal'], vertical=a['vertical'],
                                     wrap_text=a['wrap_text'], shrink_to_fit=a['shrink_to_fit'],
                                     indent=a['indent'], text_rotation=a['text_rotation'])
        except: pass
        try:
            nc.number_format = s['number_format']
        except: pass
    for cl, w in snap['col_widths'].items():
        ws.column_dimensions[cl].width = w
    for ri, h in snap['row_heights'].items():
        ws.row_dimensions[ri].height = h
    for mg in snap['merges']:
        try: ws.merge_cells(mg)
        except: pass

def generer_template_colorie(template_path, planning_classes, annee_debut, output_path):
    # Charger le template UNE SEULE FOIS et le mettre en mémoire
    wb_tpl = load_workbook(template_path)
    ws_src = wb_tpl.active
    struct = detect_structure(ws_src)
    cc = struct['class_cols']; colors = struct['class_colors']
    fdr = struct['first_data_row']; dlc = struct['day_label_col']

    snap = snapshot_template(ws_src)
    del wb_tpl  # libérer la mémoire immédiatement

    ci = 0
    for nom in cc:
        if nom not in colors:
            colors[nom] = DEFAULT_COLORS[ci % len(DEFAULT_COLORS)]; ci += 1

    mois_scolaire = [(9,annee_debut),(10,annee_debut),(11,annee_debut),(12,annee_debut),
                     (1,annee_debut+1),(2,annee_debut+1),(3,annee_debut+1),(4,annee_debut+1),
                     (5,annee_debut+1),(6,annee_debut+1),(7,annee_debut+1),(8,annee_debut+1)]
    jf_nom = {0:'Lundi',1:'Mardi',2:'Mercredi',3:'Jeudi',4:'Vendredi'}

    wb_out = Workbook(); wb_out.remove(wb_out.active)
    total = 0

    for (mois, annee) in mois_scolaire:
        ws = wb_out.create_sheet(title=f"{MOIS_FR[mois]} {annee}")
        apply_snapshot(ws, snap)  # copie depuis le dict en mémoire

        for ri in range(1, 4):
            for ci2 in range(1, snap['max_col']+1):
                v = ws.cell(row=ri, column=ci2).value
                if isinstance(v, str) and re.search(r'20\d\d', v):
                    ws.cell(row=ri, column=ci2).value = f"{MOIS_FR[mois].upper()} {annee}"

        nb_j = calendar.monthrange(annee, mois)[1]
        jours_ouvres = [datetime.date(annee, mois, j) for j in range(1, nb_j+1)
                        if datetime.date(annee, mois, j).weekday() < 5
                        and datetime.date(annee, mois, j) not in FERIES]

        row_idx = fdr
        for d in jours_ouvres:
            ds = d.strftime('%Y-%m-%d')
            ws.cell(row=row_idx, column=dlc).value = jf_nom[d.weekday()]
            ws.cell(row=row_idx, column=dlc+1).value = d.day
            for classe_nom, col_idx in cc.items():
                a_cours = any(
                    ds in pc['jours']
                    for pc in planning_classes
                    if pc['nom'] and (pc['nom'].strip().lower() in classe_nom.lower()
                                      or classe_nom.lower() in pc['nom'].strip().lower())
                )
                cell = ws.cell(row=row_idx, column=col_idx)
                if a_cours:
                    color = colors.get(classe_nom, 'FFD9D9D9')
                    if len(color) == 6: color = 'FF' + color
                    cell.fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
                    total += 1
                else:
                    cell.fill = PatternFill(fill_type=None); cell.value = None
            row_idx += 1

    wb_out.save(output_path)
    return total, len(mois_scolaire)

# ── Routes ────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generer', methods=['POST'])
def generer():
    sid = str(uuid.uuid4())[:8]; wd = os.path.join(UPLOAD_FOLDER, sid); os.makedirs(wd)
    def save(f, sub=''):
        d = os.path.join(wd, sub) if sub else wd; os.makedirs(d, exist_ok=True)
        p = os.path.join(d, f.filename); f.save(p); return p
    try:
        cf = request.files.getlist('classes'); df = request.files.getlist('dispos')
        ff = request.files.get('formateurs'); tf = request.files.get('template')
        mois = set(json.loads(request.form.get('mois', '[]')))
        if not all([cf, df, ff, tf, mois]): return jsonify({'error': 'Fichiers manquants'}), 400
        cp = [save(f, 'classes') for f in cf if f.filename]
        dp = [save(f, 'dispos')  for f in df if f.filename]
        fp = save(ff); tp = save(tf)
        pcs = [c for c in [parse_planning_classe(p) for p in cp] if c]
        dispos = [parse_disponibilite(p) for p in dp]
        aff = parse_tableau_formateurs(fp)
        assignment, stats, heures = assigner(pcs, dispos, aff)
        on = f"Planning_{'_'.join(m.replace('/','') for m in sorted(mois))}.xlsx"
        op = os.path.join(wd, on)
        ecrire_planning(tp, assignment, mois, op)
        return jsonify({
            'sessions_assignees': stats['assigned'], 'creneaux_sans_prof': stats['warn'],
            'formateurs_actifs': len(heures), 'classes': len(pcs),
            'mois': ', '.join(f"{MOIS_FR[int(m.split('/')[0])]} {m.split('/')[1]}" for m in mois),
            'fichier': on, 'session_id': sid,
            'heures_formateurs': {p: dict(m) for p, m in heures.items()}
        })
    except Exception as e:
        import traceback; return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/generer-template', methods=['POST'])
def generer_template():
    sid = str(uuid.uuid4())[:8]; wd = os.path.join(UPLOAD_FOLDER, sid); os.makedirs(wd)
    def save(f, sub=''):
        d = os.path.join(wd, sub) if sub else wd; os.makedirs(d, exist_ok=True)
        p = os.path.join(d, f.filename); f.save(p); return p
    try:
        cf = request.files.getlist('classes'); tf = request.files.get('template')
        annee = int(request.form.get('annee', 2026))
        if not cf or not tf: return jsonify({'error': 'Fichiers manquants'}), 400
        cp = [save(f, 'classes') for f in cf if f.filename]; tp = save(tf)
        pcs = [c for c in [parse_planning_classe(p) for p in cp] if c]
        on = f"Template_Affichage_{annee}_{annee+1}.xlsx"; op = os.path.join(wd, on)
        total, nb = generer_template_colorie(tp, pcs, annee, op)
        return jsonify({
            'fichier': on, 'session_id': sid, 'nb_mois': nb,
            'cases_colories': total, 'classes': len(pcs),
            'noms_classes': [c['nom'] for c in pcs if c.get('nom')]
        })
    except Exception as e:
        import traceback; return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/telecharger/<session_id>/<filename>')
def telecharger(session_id, filename):
    path = os.path.join(UPLOAD_FOLDER, session_id, filename)
    if not os.path.exists(path): return "Fichier introuvable", 404
    @after_this_request
    def cleanup(r):
        try: shutil.rmtree(os.path.join(UPLOAD_FOLDER, session_id))
        except: pass
        return r
    return send_file(path, as_attachment=True, download_name=filename)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
