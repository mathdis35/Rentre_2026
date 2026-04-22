# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## RÈGLE ABSOLUE — Fichier template Excel

**INTERDICTION TOTALE de modifier `template_vierge_SEPTEMBRE_2026.xlsx`** — ce fichier est la référence de mise en forme et ne doit jamais être écrasé, remplacé, ou modifié par Claude, même pour ré-encoder la base64. Toute mise à jour de `_TEMPLATE_VIERGE_B64` dans `app.py` doit être faite manuellement par l'utilisateur.

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

# No test suite currently exists
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
- `JOURS_COLS` — detected automatically from the embedded template: currently `[2, 26, 54, 76]` (one per section). These are the columns where day labels and numbers are written. **Do not hardcode** — they are derived by `_detect_template_constants()` at startup by finding section starts from the header row.
- `SLOT_ROWS` — `[7 + (i // 5) * 31 + (i % 5) * 6 for i in range(25)]`. Slot 0 = row 7 (Monday week 1). Each day slot = 6 rows. Each week = 31 rows (5 days × 6 + 1 separator). **Do not change** — tied to the physical template structure.
- `TEMPLATE_LAST_ROW = 160` — the closing border row of the template. **Do not change.**
- `DEFAULT_COLORS` — layout constants tied to the template structure
- `_TEMPLATE_VIERGE_B64` — the reference template embedded as base64 at line ~101. **Always re-encode after modifying `template_vierge_SEPTEMBRE_2026.xlsx`** using: `base64.b64encode(open('template_vierge_SEPTEMBRE_2026.xlsx','rb').read()).decode('ascii')`

### Template structure (ne pas modifier sans mettre à jour les constantes)

Le fichier `template_vierge_SEPTEMBRE_2026.xlsx` est la référence. Structure fixe :

- **Row 1** : titre mois (ex: `SEPTEMBRE 2026`), présent dans les 4 sections
- **Row 4** : en-têtes des classes (ex: `BAC PRO 26`, `BTS MCO 99`...)
- **Rows 7–160** : corps du tableau — 25 slots de 6 lignes chacun, 4 séparateurs de semaine
- **Row 160** : ligne de fermeture (bordure `medium` top) — **ne jamais supprimer**

**4 sections côte à côte** (séparées par des colonnes étroites ~1 char) :
| Section | Col dates | Premières classes |
|---|---|---|
| 1 (BAC PRO) | col 2 (B) | col 4 |
| 2 (BTS MCO) | col 26 (Z) | col 28 |
| 3 (BTS GPME/NDRC) | col 54 (BB) | col 56 |
| 4 (RDC/RH/M) | col 76 (BX) | col 78 |

**Colonnes séparatrices** (largeur ~1 char, grises visuellement) : cols 1, 3, 6, 9, 12, 15, 18, 21, 25, 27, 30... toutes < 2 chars de large.

**Séparateurs horizontaux entre semaines** (lignes grises) : rows 33–37, 64–68, 95–99, 126–130. Hauteur : 8.25pt (≈11px).

### Logique de génération du template vierge (`generate_month_sheet_delete`)

1. **Effacer** les dates existantes du template (issues de septembre) dans toutes les `JOURS_COLS`
2. **Supprimer les slots vides du début** : `slot_deb = premier_jour.weekday()` slots × 6 lignes depuis row 7
3. **Recalculer** `remaining_slots = [r - lignes_debut for r in SLOT_ROWS[slot_deb:]]`
4. **Écrire** le label et le numéro du jour dans chaque `JOURS_COLS` pour chaque jour ouvré
5. **Supprimer les slots vides de fin** : de `remaining_slots[nb_jours]` jusqu'à `TEMPLATE_LAST_ROW - 2 - lignes_debut`

Le `-2` dans `delete_until` est intentionnel : il préserve la ligne de fermeture (row 160 décalée) sans laisser de ligne vide résiduelle avant elle.

### Frontend compatibility

Routes support two naming conventions for uploaded files:
- **New front:** `planning_0`, `planning_1`, ... + single `disponibilites` field
- **Old front:** `classes` (getlist) + `dispos` (getlist)

Month selection is sent as `mois_json` (list of `[annee, mois]` pairs) with fallback to `annee_debut`/`mois_debut`/`annee_fin`/`mois_fin` range params.
