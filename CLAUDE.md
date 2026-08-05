# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## How to run

```bat
rodar.bat
```

Or directly:

```
python update_performance.py
```

## Dependencies

```
pip install openpyxl pandas anthropic
```

Requires Python 3.12+ and Google Drive mapped at `P:\`.

## Architecture

Single-script project (`update_performance.py`) that runs interactively each month to append one row to each tab of `Performance MBS MKTPLACE.xlsx`.

### Data flow

1. **Month detection** — scans filenames in `MATERIAIS_DIR` to infer year/month from the report filename patterns.
2. **Automatic file processing**:
   - `processar_ml_evolucao()` — aggregates daily rows from the ML Excel report.
   - `processar_shopee_overview()` — reads the single summary row from the Shopee overview Excel.
   - `processar_shopee_ads_csv()` — aggregates all ad groups from the Shopee CSV.
3. **AI image reading** — `processar_imagens_ia()` sends each `.jpg`/`.png` to Claude (`claude-haiku-4-5-20251001`) using `PROMPT_VISAO`. Claude returns a JSON identifying the image type (`ml_ads`, `ml_af`, `shopee_af`) and extracting numeric values. The user confirms or edits each result. Falls back to manual input (`_fallback_manual()`) if AI fails.
4. **Site tab** — always manual (`coletar_site()`).
5. **Spreadsheet write** — `atualizar_aba()` finds or creates the month row, copies styles and formulas from the row above (formula columns are listed in `ML_FORMULA_COLS` / `SHOPEE_FORMULA_COLS`), then writes values. `ajustar_formula()` increments relative row references when copying formulas down. A `.bkp_YYYYMM.xlsx` backup is made before saving. `update_performance.py` itself calls `wb.save(PERFORMANCE_FILE)` with no error handling — if the file is locked (open in Excel or syncing), it raises `PermissionError` and the run fails. The locked-file fallback (save to `~temp_*.xlsx` with manual replacement instructions) is only implemented in the one-off `atualizar_MMMYY.py` scripts, not in the main script.

### Key constants

- `ML_COLS` / `SHOPEE_COLS` / `SITE_COLS` — 1-based column index for every field in each tab.
- `ML_FORMULA_COLS` / `SHOPEE_FORMULA_COLS` — columns that hold Excel formulas copied from the row above (never overwritten with data).
- `ML_LINHA_DADOS` = 3, `SHOPEE_LINHA_DADOS` = 3, `SITE_LINHA_DADOS` = 4 — first data row in each tab.

### Input files location

All input files (Excel reports, CSV, screenshots) go directly into `MATERIAIS_DIR`:

```
P:\Meu Drive\Empresas\MBS Pro Grooming\Desenvolvimento de relatórios\
```

That is the same directory where `update_performance.py` lives. The README mentions a `materiais\` subfolder but the code constant `MATERIAIS_DIR` points here.

### Input file naming patterns

| Pattern | Source |
|---|---|
| `Relatorio_evolucao_negocio_YYYY_MM_DD-YYYY_MM_DD.xlsx` | Mercado Livre |
| `sales_overview_YYYYMMDD-YYYYMMDD.xlsx` | Shopee |
| `Dados+Gerais+de+Anúncios+Shopee-DD_MM_YYYY-DD_MM_YYYY.csv` | Shopee ads |
| `*.jpg` / `*.jpeg` / `*.png` | Screenshots (ML ads, ML afiliados, Shopee afiliados) |

### API key

Loaded from `config.env` in the project root (`ANTHROPIC_API_KEY=...`). The file is gitignored.

### One-off backfill scripts (`atualizar_MMMYY.py`)

When a screenshot is missing a field the AI cannot extract, create a script like `atualizar_maio26.py` that:
1. Imports constants and processing functions from `update_performance.py` (via `sys.path.insert`, since it lives in the same directory rather than being an installed package).
2. Hardcodes verified values for the fields the AI missed.
3. Calls `processar_ml_evolucao()` / `processar_shopee_overview()` / `processar_shopee_ads_csv()` normally for file-based data.
4. Merges everything and calls `atualizar_aba()` directly — bypassing the interactive flow.
5. Wraps the final `wb.save()` in `try/except PermissionError`, falling back to `~temp_*.xlsx` with printed instructions to replace the original manually once it's unlocked. Existing scripts (`atualizar_abril26.py`, `atualizar_maio26.py`, `atualizar_junho26.py`) all follow this pattern — replicate it in new ones.

These scripts are not for regular monthly use; they exist only when the standard script can't complete a month unattended.
