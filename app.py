import os, re, shutil, datetime, uuid, json
from pathlib import Path
from collections import defaultdict
from flask import Flask, render_template, request, jsonify, send_file, after_this_request
import subprocess

try:
    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment
except ImportError:
    pass

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB
UPLOAD_FOLDER = '/tmp/plannipro'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

COURS_COLOR  = 'FFA6CAF0'
MOIS_ABBR = {
    'août':8,'aout':8,'sept':9,'septembre':9,'oct':10,'octobre':10,
    'nov':11,'novembre':11,'déc':12,'dec':12,'décembre':12,
    'janv':1,'jan':1,'janvier':1,'févr':2,'fev':2,'février':2,
    'mars':3,'avr':4,'avril':4,'mai':5,'juin':6,'juil':7,'juillet':7
}
MOIS_FR = ['','Janvier','Février','Mars','Avril','Mai','Juin',
           'Juillet','Août','Septembre','Octobre','Novembre','Décembre']

# ════════════════════════════════════════════
def xls_to_xlsx(filepath):
    """Convertit .xls → .xlsx via libreoffice"""
    fp = Path(filepath)
    if fp.suffix.lower() != '.xls':
        return str(fp)
    out_dir = fp.parent
    result = subprocess.run(
        ['libreoffice', '--headless', '--convert-to', 'xlsx', str(fp), '--outdir', str(out_dir)],
        capture_output=True, text=True, timeout=30
    )
    xlsx_path = out_dir / (fp.stem + '.xlsx')
    return str(xlsx_path) if xlsx_path.exists() else str(fp)

def parse_planning_classe(filepath):
    """Retourne {'nom': str, 'jours': [date_str, ...]}"""
    fp = xls_to_xlsx(filepath)
    try:
        wb = load_workbook(fp, data_only=True)
    except Exception as e:
        return {'nom': Path(filepath).stem, 'jours': [], 'error': str(e)}

    ws = wb.active
    nom_classe = None
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=col).value
        if v and isinstance(v, str) and len(v.strip()) > 3:
            nom_classe = v.strip()
            break

    jours_cours = []
    # Détecter les blocs de mois (ligne 5)
    month_blocks = []
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=5, column=col).value
        if isinstance(v, str):
            v_low = v.strip().lower()
            for m_name, m_num in MOIS_ABBR.items():
                if m_name in v_low:
                    yr = re.search(r'(20\d\d)', v_low)
                    year = int(yr.group(1)) if yr else None
                    month_blocks.append({'col_jour': col+1, 'col_date': col+2, 'year': year, 'month': m_num})
                    break

    for block in month_blocks:
        if not block['year']:
            continue
        cj, cd, yr, mn = block['col_jour'], block['col_date'], block['year'], block['month']
        for row in range(6, ws.max_row + 1):
            cell_jour = ws.cell(row=row, column=cj)
            bg = None
            if cell_jour.fill and cell_jour.fill.fgColor and cell_jour.fill.fgColor.type == 'rgb':
                bg = cell_jour.fill.fgColor.rgb
            if bg != COURS_COLOR:
                continue
            date_val = ws.cell(row=row, column=cd).value
            day_num = None
            if isinstance(date_val, (int, float)):
                day_num = int(date_val)
            elif isinstance(date_val, datetime.datetime):
                day_num = date_val.day
            elif isinstance(date_val, str) and date_val.strip().isdigit():
                day_num = int(date_val.strip())
            if not day_num or not (1 <= day_num <= 31):
                continue
            try:
                d = datetime.date(yr, mn, day_num)
                if d.weekday() < 5:
                    jours_cours.append(d.strftime('%Y-%m-%d'))
            except ValueError:
                pass

    return {'nom': nom_classe or Path(filepath).stem, 'jours': sorted(set(jours_cours))}

def parse_disponibilite(filepath):
    """Retourne {'nom': str, 'dispo': {date_str: {matin: bool, pm: bool}}}"""
    try:
        df = pd.read_excel(filepath, sheet_name=0, header=None)
    except Exception as e:
        return {'nom': Path(filepath).stem, 'dispo': {}, 'error': str(e)}

    nom = Path(filepath).stem
    dispo = {}

    def is_av(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return False
        return str(v).strip().upper() in ['X', '4', '✓', 'OUI', 'O']

    header_row = df.iloc[0]
    month_cols = []
    for ci, val in enumerate(header_row):
        if isinstance(val, str):
            v_low = val.strip().lower()
            for m_name, m_num in MOIS_ABBR.items():
                if m_name in v_low:
                    yr = re.search(r'(20\d\d)', v_low)
                    year = int(yr.group(1)) if yr else None
                    month_cols.append({'ci': ci, 'month': m_num, 'year': year})
                    break

    jours_ignores = ['sam', 'dim', 'férié', 'ferie', 'férie', 'fériés']

    for mc in month_cols:
        ci, mn, yr = mc['ci'], mc['month'], mc['year']
        if not yr:
            continue
        col_matin, col_pm = ci, ci + 1

        for ri in range(2, len(df)):
            row = df.iloc[ri]
            day_abbr = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
            if day_abbr in jours_ignores:
                continue
            day_num_raw = row.iloc[1]
            day_num = None
            if isinstance(day_num_raw, (int, float)) and not pd.isna(day_num_raw):
                day_num = int(day_num_raw)
            if not day_num or not (1 <= day_num <= 31):
                continue
            try:
                d = datetime.date(yr, mn, day_num)
                if d.weekday() >= 5:
                    continue
                ds = d.strftime('%Y-%m-%d')
                matin_val = row.iloc[col_matin] if col_matin < len(row) else None
                pm_val    = row.iloc[col_pm]    if col_pm    < len(row) else None
                if str(matin_val).strip().lower() in jours_ignores:
                    continue
                dispo[ds] = {'matin': is_av(matin_val), 'pm': is_av(pm_val)}
            except ValueError:
                pass

    return {'nom': nom, 'dispo': dispo}

def parse_tableau_formateurs(filepath):
    """Retourne {classe: [{formateur, matiere, heures, heures_faites}]}"""
    fp = xls_to_xlsx(filepath)
    try:
        wb = load_workbook(fp, data_only=True)
    except Exception as e:
        return {}

    assignments = defaultdict(list)
    sheet_name = next((s for s in wb.sheetnames if 'mois' in s.lower()), wb.sheetnames[-1])
    ws = wb[sheet_name]

    current_class = None
    for row in ws.iter_rows(values_only=True):
        v0 = str(row[0]).strip() if row[0] else ''
        if re.match(r'(BTS|BAC|CGC|NDRC|GPME|Master|RDC|RH|EC)\s', v0, re.I):
            current_class = v0
        elif current_class and v0 and len(row) > 1 and row[1]:
            matiere = str(row[1]).strip()
            heures_raw = row[-1]
            try:
                heures = float(str(heures_raw).split('+')[0].strip())
            except:
                heures = 0
            if heures > 0 and matiere and matiere.lower() not in ['nan', 'total']:
                assignments[current_class].append({
                    'formateur': v0, 'matiere': matiere,
                    'heures': heures, 'heures_faites': 0
                })

    return dict(assignments)

def assigner(planning_classes, dispos_formateurs, affectations):
    """Moteur d'assignation principal"""
    dispo_idx = {d['nom']: d['dispo'] for d in dispos_formateurs}
    result = defaultdict(dict)
    stats = {'assigned': 0, 'warn': 0, 'no_aff': 0}

    for classe_info in planning_classes:
        classe_nom  = classe_info['nom']
        jours_cours = classe_info['jours']

        # Trouver les affectations pour cette classe
        aff = None
        for key in affectations:
            if key.strip().lower() in classe_nom.lower() or classe_nom.strip().lower() in key.lower():
                aff = affectations[key]
                break

        if not aff:
            stats['no_aff'] += len(jours_cours)
            for jour in jours_cours:
                result[jour][classe_nom] = {'formateur': '?', 'matiere': '?', 'slot': 'matin'}
            continue

        slots = ['matin', 'pm']
        slot_idx = 0

        for jour in jours_cours:
            slot = slots[slot_idx % 2]
            slot_idx += 1

            assigned = None
            for entry in sorted(aff, key=lambda x: x['heures_faites']):
                prof_nom = entry['formateur']
                prof_dispo = dispo_idx.get(prof_nom, {})
                jour_dispo = prof_dispo.get(jour, {})

                remaining = entry['heures'] - entry['heures_faites']
                if remaining <= 0:
                    continue

                if not prof_dispo:
                    assigned = entry
                    break
                if jour_dispo.get(slot, False):
                    assigned = entry
                    break

            if assigned:
                assigned['heures_faites'] += 4
                result[jour][classe_nom] = {
                    'formateur': assigned['formateur'],
                    'matiere':   assigned['matiere'],
                    'slot':      slot
                }
                stats['assigned'] += 1
            else:
                result[jour][classe_nom] = {'formateur': '⚠️', 'matiere': '', 'slot': slot}
                stats['warn'] += 1

    # Résumé des heures
    heures_par_formateur = defaultdict(lambda: defaultdict(int))
    for day_data in result.values():
        for classe, info in day_data.items():
            if info['formateur'] not in ['?', '⚠️']:
                heures_par_formateur[info['formateur']][info['matiere']] += 4

    return dict(result), stats, dict(heures_par_formateur)

def ecrire_planning(template_path, assignment, mois_cibles, output_path):
    """Remplit le template et sauvegarde"""
    shutil.copy(template_path, output_path)
    wb = load_workbook(output_path)
    filled_total = 0

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        titre = ''
        for row in ws.iter_rows(min_row=1, max_row=3, values_only=True):
            for cell in row:
                if cell and isinstance(cell, str) and len(cell) > 3:
                    titre = cell.upper()
                    break

        month_num = None; year = None
        for m_name, m_num in MOIS_ABBR.items():
            if m_name.upper() in titre:
                yr = re.search(r'(20\d\d)', titre)
                if yr:
                    year = int(yr.group(1))
                    month_num = m_num
                    break

        if not month_num or f'{month_num}/{year}' not in mois_cibles:
            continue

        # Trouver les colonnes de classes (ligne 2-4)
        class_cols = {}
        for ri in range(1, 6):
            for ci in range(1, ws.max_column + 1):
                v = ws.cell(row=ri, column=ci).value
                if isinstance(v, str):
                    v = v.strip()
                    if any(k in v for k in ['BTS', 'BAC', 'EC ', 'EC\t', 'CGC', 'NDRC', 'GPME', 'RDC', 'RH', 'Master']):
                        class_cols[v] = ci

        # Trouver les lignes de jours
        jours_fr = {
            'lundi':0,'mardi':1,'mercredi':2,'jeudi':3,'vendredi':4,
            'lun':0,'mar':1,'mer':2,'jeu':3,'ven':4
        }
        day_rows = {}
        for ri in range(1, ws.max_row + 1):
            for ci in range(1, 4):
                v = ws.cell(row=ri, column=ci).value
                if isinstance(v, str) and v.strip().lower() in jours_fr:
                    # Chercher le numéro du jour dans la ligne suivante ou colonne +1
                    for look_ri in [ri, ri+1]:
                        for look_ci in range(1, 4):
                            dv = ws.cell(row=look_ri, column=look_ci).value
                            if isinstance(dv, (int, float)) and 1 <= int(dv) <= 31:
                                try:
                                    d = datetime.date(year, month_num, int(dv))
                                    ds = d.strftime('%Y-%m-%d')
                                    if ds not in day_rows:
                                        day_rows[ds] = ri
                                except ValueError:
                                    pass

        # Remplir
        for ds, row_idx in day_rows.items():
            if ds not in assignment:
                continue
            for classe_nom, col_idx in class_cols.items():
                for key, val in assignment[ds].items():
                    if key.strip().lower() in classe_nom.lower() or classe_nom.lower() in key.strip().lower():
                        cell = ws.cell(row=row_idx, column=col_idx)
                        if val['formateur'] == '⚠️':
                            cell.value = '⚠️'
                        else:
                            cell.value = f"{val['formateur']}\n{val['matiere']}"
                        old = cell.font
                        cell.font = Font(name=old.name or 'Calibri', size=old.size or 9, bold=True, color=old.color)
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        filled_total += 1
                        break

    wb.save(output_path)
    return filled_total

# ════════════════════════════════════════════
#  ROUTES FLASK
# ════════════════════════════════════════════
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generer', methods=['POST'])
def generer():
    session_id = str(uuid.uuid4())[:8]
    work_dir   = os.path.join(UPLOAD_FOLDER, session_id)
    os.makedirs(work_dir)

    def save(f, subfolder=''):
        d = os.path.join(work_dir, subfolder) if subfolder else work_dir
        os.makedirs(d, exist_ok=True)
        path = os.path.join(d, f.filename)
        f.save(path)
        return path

    try:
        # Récupérer les fichiers
        classes_files    = request.files.getlist('classes')
        dispos_files     = request.files.getlist('dispos')
        formateurs_file  = request.files.get('formateurs')
        template_file    = request.files.get('template')
        mois_json        = request.form.get('mois', '[]')
        mois_cibles      = set(json.loads(mois_json))

        if not all([classes_files, dispos_files, formateurs_file, template_file, mois_cibles]):
            return jsonify({'error': 'Fichiers manquants'}), 400

        classes_paths    = [save(f, 'classes')    for f in classes_files    if f.filename]
        dispos_paths     = [save(f, 'dispos')     for f in dispos_files     if f.filename]
        formateurs_path  = save(formateurs_file)
        template_path    = save(template_file)

        # 1. Plannings classes
        planning_classes = [parse_planning_classe(p) for p in classes_paths]
        planning_classes = [c for c in planning_classes if c]

        # 2. Disponibilités
        dispos = [parse_disponibilite(p) for p in dispos_paths]

        # 3. Tableau formateurs
        affectations = parse_tableau_formateurs(formateurs_path)

        # 4. Assignation
        assignment, stats, heures = assigner(planning_classes, dispos, affectations)

        # 5. Génération fichier
        mois_labels = [f"{MOIS_FR[int(m.split('/')[0])]} {m.split('/')[1]}" for m in mois_cibles]
        output_name = f"Planning_{'_'.join(m.replace('/','') for m in sorted(mois_cibles))}.xlsx"
        output_path = os.path.join(work_dir, output_name)

        filled = ecrire_planning(template_path, assignment, mois_cibles, output_path)

        # Rapport
        rapport = {
            'sessions_assignees':  stats['assigned'],
            'creneaux_sans_prof':  stats['warn'],
            'formateurs_actifs':   len(heures),
            'classes':             len(planning_classes),
            'mois':                ', '.join(mois_labels),
            'fichier':             output_name,
            'session_id':          session_id,
            'heures_formateurs':   {
                prof: {mat: h for mat, h in mats.items()}
                for prof, mats in heures.items()
            }
        }
        return jsonify(rapport)

    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/telecharger/<session_id>/<filename>')
def telecharger(session_id, filename):
    path = os.path.join(UPLOAD_FOLDER, session_id, filename)
    if not os.path.exists(path):
        return "Fichier introuvable", 404

    @after_this_request
    def cleanup(response):
        try:
            shutil.rmtree(os.path.join(UPLOAD_FOLDER, session_id))
        except:
            pass
        return response

    return send_file(path, as_attachment=True, download_name=filename)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
