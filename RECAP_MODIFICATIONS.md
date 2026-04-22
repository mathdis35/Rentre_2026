# Récap des modifications à réimplémenter dans _appliquer_mois_sur_feuille

## Situation actuelle
Le fichier `app.py` contient une **ancienne version** de `_appliquer_mois_sur_feuille` (ligne 1583).
Toutes les corrections décrites ici sont à réappliquer sur cette fonction.
La fonction actuelle a les bugs suivants : dates non écrites dans les 4 sections, 
suppression début/fin incorrecte, phantom rows, double closing line.

---

## La fonction corrigée complète

Remplacer la fonction `_appliquer_mois_sur_feuille` (ligne 1583 à ~1668) par :

```python
def _appliquer_mois_sur_feuille(ws, annee, mois):
    """
    Applique les transformations d'un mois sur une feuille déjà copiée du template.
    """
    # Titre
    nom_mois = MOIS_FR_UPPER[mois]
    for r in range(1, 4):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and 'SEPTEMBRE' in v.upper():
                ws.cell(row=r, column=c).value = v.replace(
                    "SEPTEMBRE 2026", f"{nom_mois} {annee}")

    # Tous les jours lundi-vendredi (fériés inclus — gérés par les plannings)
    jours_ouvres = []
    for j in range(1, calendar.monthrange(annee, mois)[1] + 1):
        d = datetime.date(annee, mois, j)
        if d.weekday() < 5:
            jours_ouvres.append((JOURS_FR_LIST[d.weekday()], j))

    # Suppression lignes début
    premier_du_mois = datetime.date(annee, mois, 1)
    premier = premier_du_mois
    while premier.weekday() >= 5:
        premier += datetime.timedelta(days=1)
    # slot_deb basé sur le 1er du mois (pas le 1er ouvré), plafonné à 5
    lignes_a_supprimer = min(premier_du_mois.weekday(), 5) * 6
    # Si semaine entière supprimée (1er = sam ou dim), supprimer aussi le séparateur (1 ligne)
    lignes_debut_total = lignes_a_supprimer + (1 if lignes_a_supprimer == 30 else 0)

    if lignes_debut_total > 0:
        DELETE_START = 6   # IMPORTANT : le slot lundi commence en row 6, pas 7
        DELETE_COUNT = lignes_debut_total
        FIRST_KEPT   = DELETE_START + DELETE_COUNT
        saved_heights = {r: ws.row_dimensions[r].height
                         for r in range(FIRST_KEPT, ws.max_row + 1)}
        ws.delete_rows(DELETE_START, DELETE_COUNT)
        for old_r, h in saved_heights.items():
            new_r = old_r - DELETE_COUNT
            if new_r >= DELETE_START and h is not None:
                ws.row_dimensions[new_r].height = h
        # Recréer les fusions sur les lignes décalées
        from openpyxl.worksheet.cell_range import CellRange
        target_rows = {r for r in range(DELETE_START, ws.max_row + 1)
                       if (ws.row_dimensions[r].height is None or ws.row_dimensions[r].height >= 10)}
        for mg in list(ws.merged_cells.ranges):
            if mg.min_row == mg.max_row and mg.min_row in target_rows:
                try: ws.merged_cells.ranges.discard(mg)
                except Exception: pass
        for r in target_rows:
            for c1, c2 in ALL_MERGE_PAIRS:
                ws.merged_cells.ranges.add(
                    CellRange(f"{get_column_letter(c1)}{r}:{get_column_letter(c2)}{r}"))

    # Positions des slots après suppression début
    # IMPORTANT : calculer depuis SLOT_ROWS, jamais scanner les cellules existantes
    first_surviving = 6 + lignes_debut_total
    surviving_slots = [r for r in SLOT_ROWS if r >= first_surviving]
    jours_pos = [(r - lignes_debut_total, r - lignes_debut_total + 1)
                 for r in surviving_slots]

    # Écrire les labels dans les 4 sections (JOURS_COLS = [2, 26, 54, 76])
    for i, (rl, rn) in enumerate(jours_pos):
        if i < len(jours_ouvres):
            label, num = jours_ouvres[i]
            for col in JOURS_COLS:
                ws.cell(row=rl, column=col).value = label
                ws.cell(row=rn, column=col).value = num
        else:
            for col in JOURS_COLS:
                ws.cell(row=rl, column=col).value = None
                ws.cell(row=rn, column=col).value = None

    # Suppression slots vides en fin
    nb_slots_fin = len(jours_pos) - len(jours_ouvres)
    if nb_slots_fin > 0:
        first_empty_template_row = surviving_slots[len(jours_ouvres)]
        last_used_template_row   = surviving_slots[len(jours_ouvres) - 1]
        gap = first_empty_template_row - last_used_template_row
        sep = gap - 6  # 0 si même semaine, 1 si semaine différente (séparateur)
        delete_from = first_empty_template_row - lignes_debut_total - sep
        ws.delete_rows(delete_from, ws.max_row - delete_from + 1)

    # closing_row = dernière ligne restante après toutes les suppressions
    closing_row = ws.max_row

    # Nettoyer row_dimensions et merge ranges fantômes
    for phantom_r in list(ws.row_dimensions.keys()):
        if phantom_r > closing_row:
            del ws.row_dimensions[phantom_r]
    for mg in list(ws.merged_cells.ranges):
        if mg.min_row > closing_row or mg.max_row > closing_row:
            ws.merged_cells.ranges.discard(mg)

    # Restaurer le style de la closing row depuis le template (h=6.6, border=medium, fill=FFC0C0C0)
    import base64 as _b64, io as _io, openpyxl as _oxl
    _tpl_ws = _oxl.load_workbook(_io.BytesIO(_b64.b64decode(_TEMPLATE_VIERGE_B64))).active
    ws.row_dimensions[closing_row].height = _tpl_ws.row_dimensions[TEMPLATE_LAST_ROW].height
    for c in range(1, _tpl_ws.max_column + 1):
        src_cell = _tpl_ws.cell(TEMPLATE_LAST_ROW, c)
        dst_cell = ws.cell(closing_row, c)
        if src_cell.has_style:
            from openpyxl.styles import Border, Side, PatternFill
            import copy
            b = src_cell.border
            dst_cell.border = Border(
                left=Side(border_style=b.left.border_style,   color=copy.copy(b.left.color)),
                right=Side(border_style=b.right.border_style, color=copy.copy(b.right.color)),
                top=Side(border_style=b.top.border_style,     color=copy.copy(b.top.color)),
                bottom=Side(border_style=b.bottom.border_style, color=copy.copy(b.bottom.color)),
            )
            fi = src_cell.fill
            dst_cell.fill = PatternFill(fill_type=fi.fill_type,
                                        fgColor=copy.copy(fi.fgColor),
                                        bgColor=copy.copy(fi.bgColor))

    return len(jours_ouvres)
```

---

## Erreurs critiques à ne pas reproduire

### 1. DELETE_START = 7 (au lieu de 6)
Le slot lundi commence en **row 6** dans le template, pas row 7.
`SLOT_ROWS[0] = 7` est la ligne du *label* jour, mais la première ligne du slot est row 6.
→ Toujours `DELETE_START = 6`

### 2. lignes_a_supprimer basé sur premier_ouvre.weekday()
Pour les mois commençant un week-end (nov, mai, août...), le premier ouvré est lundi (weekday=0)
→ 0 lignes supprimées → slot vide row 6 reste visible.
→ Toujours baser sur `premier_du_mois.weekday()`, pas le premier ouvré.

### 3. Oublier le séparateur inter-semaine au début
Quand `lignes_a_supprimer == 30` (semaine entière), une ligne de séparateur reste.
→ `lignes_debut_total = lignes_a_supprimer + (1 if lignes_a_supprimer == 30 else 0)`

### 4. jours_pos scanné depuis les cellules existantes
Scanner "Lundi", "Mardi"... dans les cellules donne des positions basées sur le contenu septembre résiduel.
→ Toujours calculer depuis `SLOT_ROWS`.

### 5. Formule naïve 6 + i*6 pour jours_pos
Ne tient pas compte des séparateurs inter-semaine (gap=7 entre vendredi et lundi suivant).
→ Utiliser `SLOT_ROWS` directement.

### 6. delete_from = last_used_rl + 6 pour la fin
Quand le dernier jour est un vendredi, `last_used_rl + 6` tombe sur le séparateur inter-semaine
→ double ligne grise en bas.
→ Calculer via `gap = first_empty - last_used` et soustraire `sep = gap - 6`.

### 7. closing_row calculé manuellement
`TEMPLATE_LAST_ROW - lignes_debut_total - nb_slots_fin * 6` ne tient pas compte des séparateurs absorbés.
→ Toujours `closing_row = ws.max_row` après toutes les suppressions.

### 8. Ne supprimer que les row_dimensions (pas delete_rows physique)
Les cellules avec style font monter `ws.max_row` même sans contenu.
→ Faire `ws.delete_rows(phantom_start, ...)` puis nettoyer row_dimensions et merged_cells.

### 9. Filtrer les fériés dans jours_ouvres
Les fériés doivent apparaître dans le template vierge.
→ `if d.weekday() < 5` sans `and d not in FERIES`.

---

## Autres modifications hors _appliquer_mois_sur_feuille

### Frontend (templates/index.html)
- Suppression de la carte "Mode de génération" (✂️ / 👁️)
- Suppression du handler JS associé
- Suppression de l'affichage du mode dans le résultat

### Backend (app.py) — non supprimé mais inutilisé
Ces fonctions existent toujours mais ne sont plus appelées depuis le frontend :
- `generate_month_sheet_hide` (ligne 993)
- `generate_month_sheet` dispatcher (ligne 1037)
- `_supprimer_lignes_masquees_xml` (ligne 1047)
- `_copier_feuille_rapide` (ligne 1256)

Elles peuvent être supprimées sans impact.

---

## Structure du template (ne pas modifier)

```
Row 6      : première ligne du slot lundi semaine 1  ← DELETE_START
Rows 6-11  : slot lundi (6 lignes)
Rows 12-17 : slot mardi
...
Rows 31-36 : slot vendredi semaine 1
Row 37     : séparateur inter-semaine (1 ligne)
Row 38     : début slot lundi semaine 2
...
Row 160    : closing row (h=6.6, border=medium, fill=FFC0C0C0)

SLOT_ROWS = [7 + (i // 5) * 31 + (i % 5) * 6 for i in range(25)]
Gap intra-semaine = 6, gap inter-semaine = 7

JOURS_COLS      = [2, 26, 54, 76]
TEMPLATE_LAST_ROW = 160
```
