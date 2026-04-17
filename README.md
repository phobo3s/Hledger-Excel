# hledger-excel

A personal finance tracker that bridges **Microsoft Excel VBA** with [hledger](https://hledger.org/) plain-text double-entry accounting. Automates bank data collection, applies rule-based transaction categorization, and exports a fully valid hledger journal — then launches `hledger-ui` for rich terminal reporting.

---

## Features

- **TEB bank automation** — fetches account transactions, investment account statements, and credit card activity directly from the CEPTETEB desktop app via the Windows Accessibility API (no scraping, no unofficial API)
- **Rule-based categorization** — define CONTAINS / EXACT / REGEX rules in a spreadsheet; the engine auto-assigns accounts and rewrites descriptions
- **Duplicate detection** — before import, checks existing ledger entries by date + amount to avoid double-posting
- **Similar-transaction matching** — surfaces past entries with the same description so you can reuse categorizations
- **hledger journal export** — generates a valid `.hledger` file with commodity lots, buy/sell/dividend/interest postings, and market price directives
- **Portfolio tracking** — FIFO cost-basis calculation for equities; supports buy, sell, dividend, interest, and withdrawal operations
- **Price data parsing** — converts Portfolio Performance CSV exports into hledger `P` price directives
- **UTF-8 source control workflow** — `VBAExporter` exports all VBA modules as UTF-8 BOM files and re-imports them, replacing the need for RubberDuck

---

## Architecture

```
hledger-excel/
├── src/                        ← VBA source modules (tracked by git)
│   ├── Config.bas              ← All paths & settings (ThisWorkbook.Path-relative)
│   ├── LogManager.bas          ← Logging → LOGS sheet + Immediate Window
│   ├── MainModule.bas          ← Core: account aggregation + hledger export
│   ├── BankGetter.bas          ← TEB automation (Accessibility API)
│   ├── Importer.bas            ← CSV import wizard with dedup & rules
│   ├── PriceParserMod.bas      ← Portfolio Performance price CSV → .hledger
│   ├── Rules.bas               ← Rule engine (reads RULES sheet)
│   ├── HledgerMode.bas         ← Interactive hledger command runner (in-sheet)
│   ├── ImportExportHelper.bas  ← Worksheet ↔ CSV bulk export/import
│   ├── VBAExporter.bas         ← UTF-8 BOM module export + import
│   ├── BigNumbersMod.bas       ← Arbitrary-precision division for cost basis
│   ├── modFindAll64.bas        ← FindAll utility (Chip Pearson, 64-bit)
│   ├── BringWindowToFront.bas  ← Win32 window handle helper
│   ├── JSON.bas                ← JSON parser (omegastripes)
│   └── TestUTF8.bas            ← Turkish character round-trip tests
├── TheBerk_Template.xlsm       ← Blank workbook template (no personal data)
├── .gitignore
└── README.md

── (gitignored) ──────────────────────────────────────────
├── Main.hledger                ← Your journal (personal data)
├── Commodity-Prices.csv        ← Price history (personal data)
└── *.csv / *.txt               ← All other data files
```

---

## Sheet Structure

| Sheet (CodeName) | Purpose |
|---|---|
| `MAINLedger` | Master transaction ledger — one transaction per 2 rows |
| `AccountList` | Chart of accounts |
| `CommodityList` | Tracked commodities / tickers |
| `IMPORT` | Staging area for CSV imports |
| `RULES` | Categorization rules table |
| `Bank_Info` | Raw data fetched from TEB |
| `LOGS` | Runtime log (written by LogManager) |
| `Hledger` | hledger command runner worksheet |
| Account sheets (green tab) | Per-account transaction registers |

### MAINLedger row layout

Each transaction occupies **2 rows**:

| Col | Header | Example |
|-----|--------|---------|
| 1 | Date | `14.02.2024` |
| 2 | Transaction Code | `!` |
| 3 | Payee\|Note | `Market XYZ` |
| 4 | Reconciliation | running balance |
| 5 | Commodity/Currency | `CURRENCY::TRY` |
| 6 | Operation | `Buy` / `Sell` |
| 8 | Full Account Name | `Varliklar:Banka:TEB` |
| 9 | Amount | `-114.99` |
| 10 | Rate/Price | `1` |
| 11 | Reconciliation link | |
| 13 | Bank Description | original bank text |

Row 1 = debit leg, Row 2 = credit leg.

---

## Requirements

| Requirement | Notes |
|---|---|
| Windows 10/11 | Accessibility API is Windows-only |
| Microsoft Excel | `.xlsm` macro-enabled workbook |
| [hledger](https://hledger.org/install.html) | Must be on PATH |
| CEPTETEB desktop app | Only for `BankGetterTEB` automation |

---

## Setup

### 1. Clone and open

```bash
git clone https://github.com/YOUR_USERNAME/hledger-excel.git
cd hledger-excel
```

Open `TheBerk_Template.xlsm` in Excel. Rename it to whatever you like.

### 2. Trust VBA project access

`Excel Options → Trust Center → Trust Center Settings → Macro Settings`  
☑ **Trust access to the VBA project object model**

This is required for `VBAExporter` to import/export modules.

### 3. Import VBA source

In the VBA IDE (`Alt+F11`), import `src/VBAExporter.bas` manually (File → Import File).  
Then run:

```vba
VBAExporter.ImportAllModules
```

This imports all remaining modules from `src/` with correct UTF-8 encoding.

### 4. Configure your accounts

- Rename the green-tab sheets to match your real account names
- Fill in the `RULES` sheet with your categorization patterns
- Fill in `AccountList` and `CommodityList`

---

## Workflow

### Day-to-day import

1. Export a CSV from your bank (or run `BankGetterTEB` for TEB auto-fetch)
2. Paste into the `IMPORT` sheet, set the target account sheet name in cell `A2`
3. Run `Importer.ImporterBegin` — rules fire, duplicates are flagged, similar entries are surfaced
4. Review and confirm

### Generate hledger journal

```vba
CreateAllFilesAKATornado   ' aggregates all accounts → exports Main.hledger → launches hledger-ui
```

Or step-by-step:

```vba
ExportHledgerFile          ' just export, no UI launch
```

### Price data

Export from [Portfolio Performance](https://www.portfolio-performance.info/) as CSV, then:

```vba
PriceParserMod.ParsePricesFrom_PortfolioPerformance
```

Generates a `.hledger` price file alongside the CSV.

### Sync VBA source ↔ Excel

```vba
VBAExporter.ExportAllModulesUTF8   ' Excel → src/  (before git commit)
VBAExporter.ImportAllModules       ' src/ → Excel  (after git pull)
```

---

## hledger-ui

After running `CreateAllFilesAKATornado`, a terminal opens with:

```
hledger-ui -f ./Main.hledger -w -3 -X TRY --infer-market-prices -E --theme=terminal
```

Balances are converted to TRY using inferred market prices from your price directives.

---

## RULES Sheet Format

| Col | Field | Description |
|-----|-------|-------------|
| A | Active | `TRUE` / `FALSE` |
| B | DescRuleType | `CONTAINS`, `EXACT`, or `REGEX` |
| C | Description | Pattern to match against transaction description |
| D | AmountOp | `=`, `>=`, `>`, `<=`, `<`, or blank (match any) |
| E | Amount | Amount threshold |
| F | Account | Source account filter (blank = any) |
| G | ToAccount | hledger account to assign |
| H | NewDescription | Rewritten description (blank = keep original) |
| I | Special | `Buy/Sell` for commodity transactions, else transfer type |
| J | Priority | Higher number wins when multiple rules match |

---

## Encoding Notes

Turkish characters (`İ ı Ş ş Ğ ğ Ç ç Ö ö Ü ü`) require careful handling across the VBA ↔ file boundary:

- All VBA source files are stored as **UTF-8 BOM**
- Turkish string literals in `.bas` files use `ChrW()` to avoid encoding corruption:
  ```vba
  ChrW(304) = İ,  ChrW(305) = ı
  ChrW(350) = Ş,  ChrW(351) = ş
  ChrW(286) = Ğ,  ChrW(287) = ğ
  ```
- `VBAExporter` handles the ANSI ↔ UTF-8 BOM re-encoding automatically

---

## License

MIT — see [LICENSE](LICENSE)
