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

1. **Month detection** — scans filenames in `P:\Meu Drive\Empresas\MBS Pro Grooming\materiais\` to infer year/month.
2. **Automatic file processing**:
   - `processar_ml_evolucao()` — aggregates daily rows from the ML Excel report.
   - `processar_shopee_overview()` — reads the single summary row from Shopee overview Excel.
   - `processar_shopee_ads_csv()` — aggregates all ad groups from the Shopee CSV.
3. **AI image reading** — `processar_imagens_ia()` sends each `.jpg`/`.png` to Claude (`claude-haiku-4-5-20251001`) using `PROMPT_VISAO`. Claude returns a JSON identifying the image type (`ml_ads`, `ml_af`, `shopee_af`) and extracting numeric values. The user confirms or edits each result. Falls back to manual input if AI fails.
4. **Site tab** — always manual (`coletar_site()`).
5. **Spreadsheet write** — `atualizar_aba()` finds or creates the month row, copies styles/formulas from the row above (formula columns are listed in `ML_FORMULA_COLS` / `SHOPEE_FORMULA_COLS`), and writes values. A `.bkp_YYYYMM.xlsx` backup is made before saving.

### Key constants

- `ML_COLS` / `SHOPEE_COLS` / `SITE_COLS` — 1-based column index for every field in each tab.
- `ML_FORMULA_COLS` / `SHOPEE_FORMULA_COLS` — columns that hold Excel formulas copied from the row above (never overwritten with data).
- `ML_LINHA_DADOS` = 3, `SHOPEE_LINHA_DADOS` = 3, `SITE_LINHA_DADOS` = 4 — first data row in each tab.

### API key

Loaded from `config.env` in the project root (`ANTHROPIC_API_KEY=...`). The file is gitignored.

### Input file naming patterns (materiais folder)

| Pattern | Source |
|---|---|
| `Relatorio_evolucao_negocio_YYYY_MM_DD-YYYY_MM_DD.xlsx` | Mercado Livre |
| `sales_overview_YYYYMMDD-YYYYMMDD.xlsx` | Shopee |
| `Dados+Gerais+de+Anúncios+Shopee-DD_MM_YYYY-DD_MM_YYYY.csv` | Shopee ads |
| `*.jpg` / `*.jpeg` / `*.png` | Screenshots (ML ads, ML afiliados, Shopee afiliados) |

### Script `atualizar_abril26.py`

One-off script used to backfill April/2026 data — not for regular use.
