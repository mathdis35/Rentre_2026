# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

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
- `JOURS_COLS`, `DEFAULT_COLORS` — layout constants tied to the template structure

### Frontend compatibility

Routes support two naming conventions for uploaded files:
- **New front:** `planning_0`, `planning_1`, ... + single `disponibilites` field
- **Old front:** `classes` (getlist) + `dispos` (getlist)

Month selection is sent as `mois_json` (list of `[annee, mois]` pairs) with fallback to `annee_debut`/`mois_debut`/`annee_fin`/`mois_fin` range params.
