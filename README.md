# SalesReport Auto

Turn structured CSV/XLSX sales exports into an Excel report of monthly sales,
branch performance and product rankings, with a record of accepted and excluded
data.

**Version 0.1.0: portfolio release candidate. Python 3.12+.** Reporting runs
locally and offline after installation, without Excel automation or a SheetMerge
installation. See the [release audit and coverage matrix](docs/RELEASE_AUDIT.md)
for verification evidence and known limits.

## Example report

The following screenshot shows the `Summary` worksheet generated from the included synthetic transaction dataset.

![SalesReport Auto transaction summary](docs/screenshots/transactions-summary.png)

The demonstration processes seven accepted transactions and reports **USD 435.75** in net revenue.

The Excel workbook also includes monthly sales, branch performance, data-quality details, and original input records.

For product-level reporting, the application supports a separate mode with product rankings and quantity metrics.

To reproduce the demonstrated transaction report:

```powershell
.\.venv\Scripts\salesreport-auto.exe --mode transactions --input demo/sales.csv --output reports/sales.xlsx
```

Use a different output filename if `reports/sales.xlsx` already exists.

All example records are synthetic. The screenshot illustrates generated output and does not represent an actual client engagement.

## Business value

Sales exports often need repeated grouping, deduplication and spreadsheet checks
before someone can compare branches or review a month. SalesReport Auto applies
one documented set of rules and produces the metrics and exceptions together.

- **Compare performance:** signed revenue, transaction counts, branch-scoped
  customer counts and average transaction values by currency, month and branch.
- **Review products:** revenue rankings, quantities and observed unit prices from
  explicitly supplied product lines.
- **Account for exclusions:** accepted, rejected and redundant occurrences,
  structured reasons, source references and monetary coverage checks.

The synthetic demos below produce USD `435.75` from 7 transactions and USD
`195.00` from 5 product lines across 4 transactions. These are reproducible
synthetic results that can be checked with the commands below.

| Mode | Input grain | Business key | Reporting |
| --- | --- | --- | --- |
| `transactions` | One complete transaction per row | `(branch_id, transaction_id)` | Currency, monthly and branch metrics |
| `products` | One product line per row | `(branch_id, transaction_id, line_id)` | The same metrics plus product rankings, quantities and prices |

**Metrics include accepted records only. Precise totals do not imply that all
source records were accepted.** A usable report may still require review.

## Quick start (Windows)

Download or clone this repository, install Python 3.12 or later, and open
PowerShell in the repository root. This example selects Python 3.12 through the
Windows `py` launcher; select your installed newer version if appropriate.

```powershell
py -3.12 -m venv .venv
.venv\Scripts\python.exe -m pip install .
.venv\Scripts\salesreport-auto.exe --help
.venv\Scripts\python.exe -m salesreport_auto --version
New-Item -ItemType Directory -Path reports -Force | Out-Null
.venv\Scripts\salesreport-auto.exe --mode transactions --input demo/sales.csv --output reports/sales.xlsx
.venv\Scripts\salesreport-auto.exe --mode products --input demo/products.csv --output reports/products.xlsx
```

Activation is optional with these explicit paths. Installation may download
build/runtime dependencies; normal reporting does not use the network. This
installs the local checkout and assumes no PyPI publication.

Open the generated `.xlsx` files in a spreadsheet viewer; generation does not
require Excel. Use a new output filename when repeating a command. Workbooks
under `reports/` are excluded from Git, so the examples are reproduced locally.

### Demo results

| Synthetic input | Generated output | Accepted records | Transactions | Net revenue | Status |
| --- | --- | --- | --- | --- | --- |
| [demo/sales.csv](demo/sales.csv) | `reports/sales.xlsx` | 7 | 7 | USD `435.75` | `completed`, exit 0 |
| [demo/products.csv](demo/products.csv) | `reports/products.xlsx` | 5 lines | 4 | USD `195.00` | `completed`, exit 0 |

The product workbook ranks Keyboard (`P001`, USD `120.00`, net quantity 4)
ahead of Mouse (`P002`, USD `75.00`, net quantity 3). Both demos preserve branch
IDs such as `001` and have no rejected or redundant occurrences.

For your own compatible files (these input filenames are placeholders):

```powershell
.venv\Scripts\python.exe -m salesreport_auto --mode transactions --input master.xlsx --output reports/master-report.xlsx
.venv\Scripts\salesreport-auto.exe --mode products --input product-lines.xlsx --output reports/product-lines.xlsx
.venv\Scripts\salesreport-auto.exe --mode transactions --input european-sales.csv --output reports/europe.xlsx --currency-precision EUR=2
```

The installed `salesreport-auto` command and `python -m salesreport_auto` share
one CLI. `--mode`, `--input` and `--output` are required for a report; `--help`
and `--version` need no input. Modes are never inferred from filenames or columns.

## Excel deliverable

Transaction reports have five visible sheets; product reports have six, in this
fixed order:

| Worksheet | Transaction order | Product order | Contents and charts |
| --- | --- | --- | --- |
| Summary | 1 | 1 | Accepted metrics per currency; status, input digest, precision, provenance and local/upstream scope |
| Monthly Sales | 2 | 2 | Observed months; a separate revenue line chart per currency |
| Branches | 3 | 3 | Branch metrics; a separate revenue column chart per currency |
| Products | - | 4 | Product metrics and rankings; a separate horizontal top-ten product chart per currency |
| Data Quality | 4 | 5 | Counts, known-amount ledgers, coverage, exclusions, warnings, review reasons and worksheet selection |
| Raw Data | 5 | 6 | Every local input occurrence, original values, classifications, diagnostics, duplicate references and derived fields |

Tables have styled headers, filters and frozen panes. Authoritative monetary
values are fixed-precision **Excel text**, calculated by Python. IDs, dates and
raw strings are literal text too, including `=1+1`, `+001`, `NA` and leading
zeroes. Use the application's totals rather than Excel SUM over text money cells.

Charts use labelled integer minor-unit helpers and currency/scale axis captions.
If a plotted absolute value exceeds 999,999,999,999,999 minor units, that chart
is omitted with `CHART_PRECISION_LIMIT` and review status; exact table totals
remain. Empty accepted results retain applicable headers and quality details,
without invented currencies or charts.

## Input schemas

Import one UTF-8 CSV (optional BOM, comma delimiter, quoted newlines supported)
or one text-cell `.xlsx` file. Headers are exact and case-sensitive, may appear
in any order, and must occur once. Blank, duplicate, missing required or unknown
headers fail the file. Optional columns may be omitted.

### Transactions

```csv
branch_id,transaction_id,transaction_date,customer_id,amount,currency
001,T001,2026-01-10,C001,100.00,USD
```

These six columns are required. Optional: `source_file`, `source_sheet`,
`source_row`. `amount` means the signed complete transaction amount. Transaction
XLSX selects exact `Master` when present, otherwise requires one worksheet;
ignored sheets are listed in Summary and Data Quality. SheetMerge masters use
this schema as-is, including provenance.

### Products

```csv
branch_id,transaction_id,transaction_date,customer_id,amount,currency,line_id,product_id,quantity,unit_price
001,T100,2026-01-10,C001,60.00,USD,L1,P001,2,30.00
001,T100,2026-01-10,C001,25.00,USD,L2,P002,1,25.00
```

These ten columns are required. Optional: `product_name`, `source_file`,
`source_sheet`, `source_row`. Product XLSX requires exactly one worksheet, even
if one is named `Master`. SheetMerge does not supply product fields.

Here `amount` means the signed **line amount** and must equal quantity multiplied
by unit price exactly. Quantities are signed whole units, with absolute value at
most 1,000,000,000; prices must be nonnegative. Zero prices and quantities are
valid; zero quantity requires zero amount. Lines within one branch-scoped
transaction must share a date, customer and currency.

### Shared value rules

- XLSX business cells must be stored as **text**, including dates, amounts,
  quantities and prices. Numeric/date/boolean/formula/error business cells are
  rejected. Optional `source_row` alone allows an exact positive numeric integer.
  Rejected decoded values/kinds are retained; original numeric/date XML spelling
  is not promised.
- Identifiers retain case, Unicode, punctuation, whitespace and leading zeroes;
  empty/whitespace-only identifiers are invalid. Dates must be valid `YYYY-MM-DD`
  text. Optional product names are labels of at most 200 characters.
- Money is plain signed decimal text, at most 512 characters, without grouping,
  symbols, exponents or surrounding whitespace. Nonzero row amounts/unit prices
  are limited to 38 digits in integer minor units. Values requiring rounding to
  fit the currency precision are rejected.
- Standalone precision defaults: `USD=2`, `PLN=2`, `JPY=0`. Repeat
  `--currency-precision CODE=N` for different uppercase three-letter codes,
  with `N` from 0 to 9, to extend/override defaults. Repeated declarations fail;
  unknown currencies are rejected and unquantified until configured.

## Exact financial calculations

Monetary tokens are explicitly parsed with `Decimal` and converted exactly to
Python integer minor units. Sums and price/amount checks use integers, without
binary floating-point business arithmetic. Aggregates can exceed the individual
row limit. Currencies remain separate in tables and charts; no exchange-rate
conversion or combined currency total is produced.

Net revenue is the signed sum of accepted amounts. Negative values reduce it;
zero and negative transactions remain countable. The label makes no tax,
gross/net accounting, settlement or revenue-recognition claim and infers no
refund link. Transaction counts use distinct `(branch_id, transaction_id)` keys.
Product lines first form transaction totals, so average transaction value divides
revenue by transactions, not line count. Customer counts use distinct
`(branch_id, customer_id)` pairs. Monthly grouping uses observed calendar months.

Averages round once for display with `ROUND_HALF_UP`, ties away from zero;
missing denominators display `N/A`. Product metrics include positive quantity,
negative quantity magnitude, net quantity, signed revenue, minimum/maximum prices
and a positive-quantity-weighted unit price. The weighted price divides amounts
on positive-quantity lines by positive quantity. Products group by currency/product
ID across branches and rank by descending revenue, then exact product ID.
Product revenue and line counts reconcile to Summary; overlapping product
transaction/customer counts must not be summed to obtain Summary.

## Data quality and report status

Every local input occurrence is classified once as `accepted`, `rejected` or
`redundant`. Counts and known amounts are checked for conservation. A bad date or
price does not erase independently parseable money; unquantified amounts remain
explicit counts, never fabricated zeroes. Structured issues/warnings, local
import locations and optional source references are preserved.

Transaction duplicates compare normalized business values, excluding provenance.
Exact copies retain the earliest physical representative; later copies are
redundant. Conflicting or invalid same-key candidates reject the whole readable
key group. In product mode, distinct line IDs stay separate, even for the same
product. An invalid/conflicting line or inconsistent transaction metadata
quarantines **every occurrence of the identified transaction**, with no accepted
or redundant remainder. Unreadable transaction keys are rejected independently.
Product label disagreements can require review without changing financial totals.

Missing all optional provenance is informational; partial/invalid provenance
warns without rejecting otherwise valid sales. Without `--bundle`, upstream
acceptance and monetary completeness remain **unknown / not supplied**, even for
`master.csv`; complete-looking source references are still unverified.

| Exit | Published status | Meaning |
| --- | --- | --- |
| 0 | `completed` | No review reasons; redundancy alone is permitted |
| 1 | `review_required` | Published with rejections, unquantified money, no accepted records, provenance/label warnings, chart limits or upstream quality concerns |
| 2 | No success status | Argument/input/integrity/writing/verification/publication failure; no success JSON |

Successful publication prints one JSON object with `status`, `mode`, `output`,
`local_counts`, `local_unquantified_count`, `local_monetary_complete`, `upstream`,
`review_reasons` and `currencies`. Money totals are strings; standalone `upstream`
is null. Header-only/all-rejected files produce review reports. A completed report
does not prove that every real-world sale or product line was supplied.

## Optional SheetMerge bundle integration

```powershell
.venv\Scripts\salesreport-auto.exe --mode transactions --input handoff/master.csv --bundle handoff --output reports/merged.xlsx
```

`handoff` represents your existing version-1 bundle containing all seven files:
`master.csv`, `master.xlsx`, `exclusions.csv`, `audit.jsonl`, `reconciliation.csv`,
`reconciliation.json`, `manifest.json`. Select one master explicitly; the two
masters are never combined. A byte-identical selected copy outside the bundle is
allowed with the original master basename. Products plus `--bundle` fails clearly.

Verification checks manifest/tool/business-key requirements, declared SHA-256
hashes, reconciliation partitions, selected-master accepted counts/totals, source
references and exclusion/duplicate references. Invalid bundles fail without
fallback. The full `manifest.configuration.currency_precision` mapping is
authoritative; CLI precision options may only repeat identical entries.

Local master statistics and upstream original-source statistics remain separate.
Upstream exclusions retain reasons/raw values in Data Quality, rather than joining
local Raw Data or sales totals. Upstream rejections/incomplete monetary coverage
require review. Hash consistency proves neither authenticity nor business
completeness. The alternate master, audit JSONL and reconciliation CSV are
hash-checked companions, not replayed metric inputs; original sources and declared
configuration fingerprints are not independently verified.

The opt-in sibling fixture has 33 original occurrences: 16 accepted, 14 rejected,
3 redundant and 5 unquantified amounts. Accepted revenue is USD `141.45`, PLN
`220.35`, JPY `1290`, with review required. These assertions are separate from the
two standalone demos. Normal use never imports/executes SheetMerge or discovers
adjacent metadata automatically. See the [specification](docs/REPORTING_SPEC.md)
for the complete bundle contract.

## File safety and limits

Sources and bundle artifacts are read-only. Output must be a new `.xlsx` file in
an existing writable directory, outside any supplied bundle. Existing targets,
input aliases and paths through symlinks/junctions are refused. Before publication,
the staged workbook is reopened to check sheets, tables, metrics, quality, literal
cell types and chart sources; input/bundle snapshots are rechecked for changes.
Publication uses an atomic same-directory hard link, requiring filesystem support,
with no overwrite fallback.

If staging cleanup is denied after publication, success/review status remains
valid and stderr identifies the owned temporary leftover. Before publication,
cleanup warnings preserve the original failure. Python callers can inspect
`ReportResult.cleanup_warning`.

Limits: 20 MiB per input, 100,000 data occurrences including internal blanks,
200 MiB advertised XLSX uncompressed size. Bundle files are limited to 20 MiB,
except streamed, hash-only audit JSONL at 200 MiB. Excel cell/XML/row limits fail
rather than truncate diagnostics. These are local business-file bounds, not a
complete hostile-file sandbox.

## Verification and developer use

```powershell
.venv\Scripts\python.exe -m pip install -e ".[dev]"
.venv\Scripts\python.exe -m pytest -q
.venv\Scripts\python.exe -m ruff check src tests
.venv\Scripts\python.exe -m pip check
```

Recorded release verification on Windows/Python 3.12.15:

| Check | Result |
| --- | --- |
| Default tests | **521 passed, 4 skipped, 0 failed** |
| Opt-in SheetMerge tests | **3 passed, 0 skipped, 0 failed** |
| Tests from extracted source distribution | **521 passed, 4 skipped, 0 failed** |
| Ruff and pip check | Passed |
| Wheel/source builds and metadata/content checks | Passed |
| Installed-wheel API/module/console checks | Passed for both modes |

The four default skips are three opt-in sibling cases and one real Windows
symlink test lacking creation privileges; simulated link guards run. Default tests
create synthetic files in pytest temporary folders and require neither network
nor SheetMerge. They cover exact money/rounding, composite keys, product quarantine,
conservation, literal text, charts, malformed inputs, bundle tampering and publication
failures. Counts are recorded test evidence, not a hosted CI badge or a coverage
percentage. See the [audit](docs/RELEASE_AUDIT.md) for the matrix and remaining gaps.

If system-temp access is restricted, use a fresh workspace basetemp. Pytest may
clear its target; do not point it at a folder containing work.

```powershell
New-Item -ItemType Directory -Path .test-tmp -Force | Out-Null
$testRun = Join-Path (Get-Location) ('.test-tmp\run-' + [guid]::NewGuid().ToString('N'))
.venv\Scripts\python.exe -m pytest -q --basetemp=$testRun
```

Optional sibling integration:

```powershell
$integrationRun = Join-Path (Get-Location) ('.test-tmp\integration-' + [guid]::NewGuid().ToString('N'))
.venv\Scripts\python.exe -m pytest -q tests/test_sheetmerge_integration.py --sheetmerge-integration --basetemp=$integrationRun
```

Those cases require the compatible `../sheetmerge` fixtures and its Python 3.12+
environment with pandas/openpyxl. They generate bundles only in isolated temporary
folders, disable sibling bytecode writes, and check source sample hashes/mtimes.
They skip with a reason when unavailable.

Typed Python APIs include `analyze_transactions`, `analyze_products`,
`generate_report` and `publish_report`. Read-only analysis, from the repository root:

```python
from salesreport_auto import analyze_transactions

analysis = analyze_transactions("demo/sales.csv")
for group in analysis.aggregates.summary:
    print(group.currency, group.metrics.net_revenue, group.metrics.transaction_count)
```

The [package modules](src/salesreport_auto) separate parsing, validation,
classification, reconciliation, integer aggregation and workbook generation.
With development tools installed, build local archives using:

```powershell
.venv\Scripts\python.exe -m build
```

The isolated build may download dependencies. Wheel/source archives go to ignored
`dist/`; source archives include docs, synthetic CSVs and complete test helpers.
Rebuild after documentation changes before sharing archives. This command does
not publish a package.

## Release status and limitations

Version 0.1.0 is a release candidate for portfolio review, not production-certified
or independently financially audited. Verification exercised Windows/Python 3.12;
other permitted Python/dependency/platform combinations, actual Excel/LibreOffice
visual rendering and full-scale performance remain unverified. The audit also
records narrower boundary-test gaps and CSV quote-grammar ambiguity.

There is no native numeric XLSX business-cell inference, fractional quantity,
discount/tax model, refund matching, FX conversion, arbitrary column mapping,
multi-file/incremental import or original-source audit replay. Product identities
and whole-unit definitions must be consistent in supplied data; missing sales or
lines cannot be detected from an incomplete export.

The [reporting specification](docs/REPORTING_SPEC.md) retains its original
planning-stage introduction. Current implementation evidence and historical
wording discrepancies are documented in the [release audit](docs/RELEASE_AUDIT.md).
Pandas is a declared dependency; current parsing and financial aggregation use
standard-library records, `Decimal` parsing and Python integers.

## License status

**No license has been selected.** There is no license file or license declaration
in `pyproject.toml`. License terms and any copyright notice remain an owner
decision; this documentation assigns neither. Portfolio publication should not
claim a specific open-source or commercial license.
