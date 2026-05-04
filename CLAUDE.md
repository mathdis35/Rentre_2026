# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## RÈGLE DE TRAVAIL — Limite de modifications

**Après 3 modifications infructueuses sur une même partie du programme, STOP.** Ne pas continuer à modifier. Prendre du recul, relire le fichier de référence (`template_OCTOBRE_2026_REF.xlsx`), et reformuler le problème avant de toucher au code.

## RÈGLE ABSOLUE — Fichier template Excel

**INTERDICTION TOTALE de modifier `template_vierge_SEPTEMBRE_2026_corrige.xlsx`** — ce fichier est la référence de mise en forme et ne doit jamais être écrasé, remplacé, ou modifié par Claude, même pour ré-encoder la base64. Toute mise à jour de `_TEMPLATE_VIERGE_B64` dans `app.py` doit être faite manuellement par l'utilisateur.

## Repository

GitHub: https://github.com/mathdis35/Rentre_2026

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run locally
flask run
# or with gunicorn (mirrors production)
gunicorn app:app --bind 0.0.0.0:5000 --timeout 300 --workers 1

# Run tests
python -X utf8 test_template_vierge.py
```

## Architecture

Single-file Flask application (`app.py`) with one HTML template (`templates/index.html`).

**Upload flow:** Files are uploaded to `/tmp/plannipro/<session_id>/`, processed in memory, and the output Excel file is served via `/telecharger/<session_id>/<filename>` then deleted.

### Parsers (input)

Three types of Excel files are parsed:

| Function | Input file | Format |
|---|---|---|
| `parse_planning_classe()` | Planning de classe | `.xls` (xlrd) or `.xlsx` (openpyxl) — detects days highlighted in blue `#A6CAF0` |
| `parse_disponibilite()` | Disponibilités formateur | `.xlsx` — columns matin/PM per month |
| `_auto_parse_formateurs()` | Tableau affectations | Auto-selects v1 (matriciel) or v2 (feuille `AFFECTATIONS`) |

`parse_tableau_formateurs_v2()` is the current format: one row per assignment with columns `CLASSE`, `FORMATEUR`, `MATIERE`, `HEURES_ANNEE` (+ optional `PRIORITE`, `ACTIF`).

### Assignment engine

`assigner(planning_classes, dispos_formateurs, affectations)` — for each school day of each class, picks the available trainer with the fewest hours done so far, alternating matin/PM slots.

### Output generation

Three routes produce Excel files:

| Route | Function | Output |
|---|---|---|
| `POST /generer` | `ecrire_planning()` | Multi-sheet planning with trainer assignments |
| `POST /generer-template` | `generer_template_colorie()` | Template with class days color-coded |
| `POST /generer-template-vierge` | `generer_excel_multifeuilles()` | Blank multi-sheet template |

`/generer` internally calls `generer_excel_multifeuilles()` first to build a blank multi-sheet base, then `ecrire_planning()` writes assignments into it.

### Key constants

- `FERIES` — hardcoded public holidays for 2026-2027 (update when extending to other years)
- `ALL_MERGE_PAIRS` — merged column pairs in the reference template (do not change)
- `JOURS_COLS` — currently `[2, 18, 38, 54]` (one per section). These are the columns where day labels and numbers are written.
- `SLOT_ROWS` — `[6 + (i // 5) * 21 + (i % 5) * 4 for i in range(25)]`. Slot 0 = row 6 (1er jour semaine 1). Chaque slot = 4 lignes. Chaque semaine = 21 lignes (5 jours × 4 + 1 séparateur). **Ne pas modifier** — lié à la structure physique du template.
- `TEMPLATE_LAST_ROW = 131` — la ligne de fermeture du template (bordure `medium` top). **Ne pas modifier.**
- `DEFAULT_COLORS` — layout constants tied to the template structure
- `_TEMPLATE_VIERGE_B64` — le template de référence embarqué en base64. **Toujours ré-encoder après modification de `template_vierge_SEPTEMBRE_2026_corrige.xlsx`** avec : `base64.b64encode(open('template_vierge_SEPTEMBRE_2026_corrige.xlsx','rb').read()).decode('ascii')`

### Template structure (ne pas modifier sans mettre à jour les constantes)

Le fichier `template_vierge_SEPTEMBRE_2026_corrige.xlsx` est la référence. Structure fixe :

- **Row 1** : titre mois (ex: `SEPTEMBRE 2026`), présent dans les 4 sections
- **Row 4** : en-têtes des classes (ex: `BAC PRO 26`, `BTS MCO 99`...)
- **Rows 6–131** : corps du tableau — 25 slots de 4 lignes chacun, 5 séparateurs de semaine
- **Row 131** : ligne de fermeture (bordure `medium` top) — **ne jamais supprimer**

**Structure d'un slot (4 lignes) :**
```
ligne slot+0 : séparateur fin (h≈11.4, border top=medium)
ligne slot+1 : label du jour ("Lundi", "Mardi"…)   ← écriture du nom du jour
ligne slot+2 : numéro du jour (1, 2, 3…)            ← écriture du numéro
ligne slot+3 : contenu (h≈16.8, border bottom=medium)
```
Représente 4 créneaux de 2h = 8h par jour (2 matin + 2 après-midi).

**Structure d'une semaine (21 lignes) :**
- 5 jours × 4 lignes = 20 lignes
- 1 ligne séparateur de semaine (h=7.2, border top+bottom=medium)

**4 sections côte à côte** (séparées par des colonnes étroites ~1 char) :
| Section | Col dates | Premières classes |
|---|---|---|
| 1 (BAC PRO) | col 2 (B) | col 4 |
| 2 (BTS MCO) | col 18 | col 20 |
| 3 (BTS GPME/NDRC) | col 38 | col 40 |
| 4 (RDC/RH/M) | col 54 | col 56 |

**Séparateurs horizontaux entre semaines** : rows 26, 47, 68, 89, 110. Hauteur : 7.2pt.

### Exemple du résultat attendu : `template_OCTOBRE_2026_REF.xlsx`

Octobre 2026 commence un Jeudi (slot_deb=3). Structure observée :
- **Row 6** : séparateur début premier slot (h=11.4, border_top=medium) — toujours présent
- **Row 7** : label "Jeudi", **Row 8** : numéro 1, **Row 9** : contenu (border_bottom=medium)
- **Row 14** : séparateur semaine (h=7.2, border_top+bottom=medium) — après Vendredi
- **Row 98** : dernier séparateur de semaine = closing row (h=7.2, border_top+bottom=medium)
- `max_row` openpyxl = 109 (fantômes) mais la vraie dernière ligne utile = **98**
- La **closing row = dernier séparateur de semaine** après le dernier jour utilisé
- Colonnes blanches sur séparateurs : 16, 36, 52

### Logique de génération du template vierge (`generate_month_sheet_delete` et `_appliquer_mois_sur_feuille`)

**Ordre impératif :**
1. **Effacer** les valeurs résiduelles dans tous les SLOT_ROWS (offsets 0–3, JOURS_COLS)
2. **Écrire** les jours à partir de `SLOT_ROWS[slot_deb]` (pas depuis 0)
3. **Supprimer le bas** : de `closing_row + 1` jusqu'à `TEMPLATE_LAST_ROW - 1` (la ligne 131 se déplace sur closing_row)
4. **Supprimer le haut** : `slot_deb * 4` lignes depuis la ligne 6
5. **Purger** `ws._cells` et `ws.row_dimensions` au-delà de la closing_row finale
6. **closing_row** = séparateur de semaine qui suit le dernier jour = `SLOT_ROWS[slot_deb + nb_jours - 1] + 4` si dernier jour = vendredi, sinon le prochain séparateur de semaine dans la liste [26,47,68,89,110]

**Lignes jamais supprimées :** 1–5 (titre/en-têtes) et la ligne 131 (closing originale).

### Frontend compatibility

Routes support two naming conventions for uploaded files:
- **New front:** `planning_0`, `planning_1`, ... + single `disponibilites` field
- **Old front:** `classes` (getlist) + `dispos` (getlist)

Month selection is sent as `mois_json` (list of `[annee, mois]` pairs) with fallback to `annee_debut`/`mois_debut`/`annee_fin`/`mois_fin` range params.
