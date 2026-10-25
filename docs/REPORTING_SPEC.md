# SalesReport Auto: business and technical specification

Specification version: 1.0. Decision date: 2026-10-25. Status: proposed MVP contract; application not implemented.

## 1. Purpose, independence, and delivery boundary

SalesReport Auto is an independent Python 3.12 command-line portfolio application that turns one explicitly structured CSV or XLSX file into an auditable Excel sales report. It uses pandas for tabular organization, openpyxl for workbook I/O and charts, and pytest and Ruff for verification. Python standard-library Decimal and integers supply monetary arithmetic.

There are two explicit modes, never inferred from filenames or available columns:

| Mode | Record grain | Business use |
| --- | --- | --- |
| `transactions` | One complete transaction per branch and transaction ID | Transaction counts, customer counts, signed revenue, average transaction value, monthly and branch analysis |
| `products` | One product line per branch, transaction ID, and line ID | The same transaction-level metrics, plus product revenue rankings, quantities, and observed unit prices |

The application must run without SheetMerge installed, without a sibling repository, without a network connection, and without bundle metadata. SheetMerge support is an optional file-format adapter, not a runtime import, subprocess, shared configuration module, or hard-coded `../sheetmerge` path. The two applications have separate release cycles.

MVP delivery means the mandatory behavior in sections 2-10 and its automated acceptance tests in section 11. This document is the only deliverable at this stage; no application, datasets, or tests are created as part of specifying it.

## 2. Authoritative SheetMerge baseline

The prior read-only inspection and the following source/tests were checked again for this specification. Paths below are relative to the SalesReport Auto repository root; they are evidence, not runtime dependencies.

| Evidence | Established contract |
| --- | --- |
| `../sheetmerge/src/sheetmerge/models.py`: `CANONICAL_COLUMNS`, `PROVENANCE_COLUMNS` | Six canonical fields followed by three provenance fields |
| `../sheetmerge/src/sheetmerge/reporting.py`: `MASTER_COLUMNS`, `_canonical_record`, `_write_csv`, `_write_workbook` | `master.csv` and `master.xlsx`; XLSX worksheet `Master`; every master value is text; ISO dates and fixed-precision amounts |
| `../sheetmerge/src/sheetmerge/reporting.py`: `_write_reports`, `export_batch` | Exact exclusion, reconciliation, audit, and manifest field names documented in section 7 |
| `../sheetmerge/src/sheetmerge/reconciliation.py`: both `to_dict` methods | Integer counts, boolean coverage flags, decimal-string amounts and integer-string minor units |
| `../sheetmerge/src/sheetmerge/consolidate.py`: `BUSINESS_KEY`, `consolidate` | Branch-scoped transaction key; redundant exact duplicates; full-group quarantine of conflicts/invalid candidates |
| `../sheetmerge/src/sheetmerge/money.py`, `config.py`, `validation.py` | Exact minor units, precision 0-9, at most 38 digits in a nonzero row amount's minor units; all six canonical fields required |
| `../sheetmerge/tests/test_reporting.py`, `test_cli.py`, `test_consolidate.py` | Literal XLSX text, conservation, exit statuses, empty masters, duplicate behavior |
| `../sheetmerge/tests/test_samples.py`, `samples/expected.json`, `samples/sources.toml` | Synthetic fixture outcomes and branch/source mappings |

SheetMerge currently reports version `0.1.0` and manifest version `1`. Its command is `sheetmerge --config TOML --input DIRECTORY --output NEW_DIRECTORY`, also available through `python -m sheetmerge`. It publishes seven files: `master.csv`, `master.xlsx`, `exclusions.csv`, `audit.jsonl`, `reconciliation.csv`, `reconciliation.json`, `manifest.json`. Its exit codes are 0 (completed), 1 (published for review), and 2 (failure). Its stdout summary is not a persisted bundle file.

SheetMerge supplies neither product IDs/names, quantities, nor unit prices. It does not define a universal gross/net/tax interpretation of `amount`. Its ingestion helper `BatchResult.to_frame()` precedes duplicate resolution and is not a final master; SalesReport Auto consumes exported masters instead.

## 3. Input transport and common value rules

### 3.1 File and header contract

- One regular `.csv` or `.xlsx` file per invocation; extension matching is case-insensitive. No directories, globbing, concatenation, automatic discovery, `.xls`, or `.xlsm` in MVP.
- CSV: strict UTF-8 decoding with an optional UTF-8 BOM, comma delimiter, double-quote escaping, LF or CRLF records, header in the first physical record. Quoted newlines are supported. `import_row` identifies a record's first physical line.
- XLSX: header in row 1. Transaction mode selects the exact `Master` worksheet if present; otherwise it requires exactly one worksheet. Product mode requires exactly one worksheet. Do not silently select the active/first sheet of an ambiguous workbook. In the `Master` case, list any unselected worksheets in Data Quality as informational metadata.
- Required headers must occur exactly once; header order is flexible. Names are exact and case-sensitive; do not trim or rename headers. Blank, duplicate, missing, or unknown headers are fatal structural errors. This intentionally catches product columns supplied in transaction mode instead of ignoring them.
- Do not allow nonempty XLSX cells beyond the defined headers. Ignore formatting-only trailing cells/rows; count blank rows within the span from row 2 through the last row with a stored value. A CSV physical blank record after the header is an accounted-for blank record, including a final extra blank line; a normal final line terminator adds no record.
- Maximum input size is 20 MiB, with at most 100,000 data records including internal blanks. XLSX ZIP members may total at most 200 MiB uncompressed; reject unsupported/encrypted archives. Reject an excessive worksheet span before expanding it. Fail rather than truncate. Bundle files have the same 20 MiB per-file budget except `audit.jsonl`, which has a 200 MiB budget and is streamed for hashing only.
- Treat these as local, trusted-business-file limits, not a complete hostile OOXML sandbox. Do not evaluate formulas, follow external links, or save source workbooks.

### 3.2 Storage types and normalization

Canonical business input values in both modes are text in CSV and must be stored as text in XLSX. Numeric/date/boolean/formula/error XLSX cells in business columns are row errors, even if their display looks correct. This narrow contract consumes SheetMerge's text masters directly and avoids recreating its numeric XML/date decoder. Users of product mode must export its documented text schema. A numeric XLSX cell is permitted only for optional `source_row` when its decoded value is an exact positive integer (not a boolean).

Read raw cells without pandas inference or NA substitution. If pandas reads CSV, require `dtype=str, keep_default_na=False, na_filter=False`; verify headers/record widths before accepting data. Never use `astype(float)`, `to_numeric` inference, `Decimal(float_value)`, or float-backed groupby/mean for business metrics.

| Logical type | Exact rule |
| --- | --- |
| Identifier | Nonempty text with at least one non-whitespace character; preserve case, Unicode, punctuation, whitespace, and leading zeroes. Literal `NA`, `NULL`, `=1+1`, and `001` are strings, not missing values or formulas. No case folding, trimming, numeric reconstruction, or Unicode normalization. Equality is exact string equality. |
| Date | Text matching `[0-9]{4}-[0-9]{2}-[0-9]{2}` and a valid Gregorian date in years 0001-9999; reject times, timezones, serials, nonpadded dates, and alternative layouts. Parse explicitly, then verify ISO round-trip equality. |
| Monetary text | At most 512 characters, matching `[+-]?[0-9]+(?:\.[0-9]+)?`; parse explicitly using `Decimal(text)`. No surrounding whitespace, grouping separators, decimal commas, exponent syntax, currency symbols, NaN, Infinity, or parentheses. Must be exactly representable at the resolved currency precision. |
| Currency | Exactly three uppercase ASCII letters, preserved as supplied. A code must have an explicitly resolved precision; no automatic uppercasing, exchange conversion, ISO lookup, or precision inference from observed amounts. |
| Quantity | Text matching `[+-]?[0-9]+`, parsed to a signed Python integer with absolute value at most 1,000,000,000. Whole units only in MVP; zero and negatives are valid. |
| Optional label | Text, up to 200 Unicode characters; blank/whitespace-only means absent. Preserve other text exactly. A nontext XLSX cell is invalid. |

Required values that are absent, empty, or whitespace-only are missing. Distinguish an actual empty value from literal strings such as `NA`. Preserve the original spelling in Raw Data; normalize accepted dates, amounts, quantity, and `source_row` only in derived fields. Canonical negative zero normalizes to zero.

### 3.3 Currency precision

Without a bundle, the explicit built-in precision table is `USD=2`, `PLN=2`, `JPY=0`. It is a limited application default, not a claim that every currency has two decimals. Optional repeatable `--currency-precision CODE=N` extends or overrides that table; `CODE` follows the currency syntax and `N` is an integer 0-9. Reject repeated declarations for the same code, even if equal. Record the effective table and which entries came from defaults/CLI.

With `--bundle`, use `manifest.configuration.currency_precision` as the complete authority, with no default-table merge. CLI declarations, if supplied, may only repeat identical bundle entries; conflicting or additional entries are fatal. This allows any configured SheetMerge currency/precision without editing its master. Unknown standalone currency codes produce rejected, unquantified rows and review, not invented precision. A user can rerun with an explicit declaration.

## 4. Mode schemas

### 4.1 Transactions

Every row describes one complete transaction. Required columns are exactly the six business columns below. The three provenance columns are optional individually; all are supplied by a SheetMerge master. No other columns are supported.

| Column | Required | Type | Meaning and validation |
| --- | --- | --- | --- |
| `branch_id` | Yes | Identifier | Branch namespace for transaction and customer IDs |
| `transaction_id` | Yes | Identifier | Transaction ID unique within its branch across the whole input; date is not part of the key |
| `transaction_date` | Yes | Date | Business calendar date of the complete transaction |
| `customer_id` | Yes | Identifier | Customer reference within the branch; anonymous/unknown customers need an explicit upstream identifier, not a fabricated value |
| `amount` | Yes | Monetary text | Signed recorded transaction amount; negatives reduce reported revenue, zeros remain countable |
| `currency` | Yes | Currency | Unit of the amount and reporting partition |
| `source_file` | No | Provenance text | Original source filename/path, retained literally and never opened or resolved |
| `source_sheet` | No | Provenance text | Original worksheet name; empty is legitimate for a CSV source |
| `source_row` | No | Provenance row | Original positive one-based physical row; decimal-digit text or the exact numeric integer exception in section 3.2 |

Provenance is not a business key and does not affect financial duplicate equality. A complete provenance reference has nonempty `source_file` and valid `source_row`; `source_sheet` can legitimately be empty. Omitted columns/empty values stay absent; never replace them with the imported master's filename/row. Always retain separate local `import_file`, `import_sheet`, and `import_row` fields so local location and upstream location remain distinguishable.

No provenance is informational and does not itself require review. Partially supplied references or invalid `source_row` values produce `INCOMPLETE_PROVENANCE` or `INVALID_PROVENANCE` warnings and review, but do not reject otherwise valid sales records. Preserve the supplied raw value. Without bundle verification, complete-looking provenance is still unverified. With a bundle, every selected master row must have a complete valid reference.

For provenance coverage, classify each local occurrence exactly once as `absent` (all three values absent), `invalid` (any supplied provenance value has the wrong storage type or a nonpositive/noninteger row), `complete` (valid file/row with optional sheet), or `partial` (the remaining cases), in that priority order. Whitespace-only provenance text counts as absent; a column header alone does not supply provenance. These counts sum to local input. Complete references are labelled unverified without a bundle; bundle agreement verifies their internal consistency only.

### 4.2 Products

Product mode is a separate line-item schema. It requires all six transaction business columns above, with **`amount` meaning this line's signed amount**, plus the following columns. The same three optional provenance columns are available. The only other optional column is `product_name`.

| Column | Required | Type | Meaning and validation |
| --- | --- | --- | --- |
| `line_id` | Yes | Identifier | Unique line identifier within `(branch_id, transaction_id)`; different lines of the same product remain different records |
| `product_id` | Yes | Identifier | Product identity, assumed consistent across branches within a currency partition; no cross-currency product consolidation |
| `quantity` | Yes | Quantity | Signed whole units on this line; negative means a quantity reversal, without inferring an original sale/refund link |
| `unit_price` | Yes | Monetary text | Nonnegative recorded amount per unit in the line currency; same exact currency precision as `amount`; zero allowed |
| `product_name` | No | Optional label | Display label only; never used as product identity |

Required business column order for documentation/export examples is `branch_id, transaction_id, transaction_date, customer_id, amount, currency, line_id, product_id, quantity, unit_price`; optional `product_name, source_file, source_sheet, source_row` follow. Input order remains flexible.

Require `amount_minor == quantity * unit_price_minor` exactly. Do not round the product, derive a missing amount, accept unexplained discounts, add tax, or distribute order-level adjustments. A zero quantity requires zero amount; a negative quantity at positive price requires a negative amount; zero price permits zero amount with any valid signed quantity. Sub-minor-unit prices, fractional quantities, discounts, bundles, and allocation rules are future schema extensions.

Within an order `(branch_id, transaction_id)`, all lines must have the same `transaction_date`, `customer_id`, and `currency`. Do not sum a repeated order-level total from each line. A SheetMerge master passed with `--mode products` fails for missing product headers; the application must never manufacture them.

Different labels on different accepted lines of the same `(currency, product_id)` produce `PRODUCT_LABEL_CONFLICT` warnings and review, without changing financial acceptance. Display the lexicographically first nonempty label, or `product_id` if none; retain every original label. This is a display policy, not product matching.

## 5. Validation, duplicates, and local quality accounting

### 5.1 Failure levels and classification order

Fatal file/run errors include missing/unreadable input, a zero-byte/wholly empty file with no header, invalid encoding, malformed CSV quoting or record width, invalid workbook/ambiguous worksheet, bad headers, exceeded limits, incompatible bundle metadata, and inability to publish a verified report. Exit 2; no final workbook.

A structurally valid header-only file produces a report with zero imported/accepted rows, no observed currency metric rows, no charts, a `NO_ACCEPTED_RECORDS` review reason, and exit 1. An all-rejected file behaves similarly but retains every local data record and reason. Do not turn missing denominators into zero averages.

Validate independent fields and money for every row before resolving keys; collect all applicable issues rather than stopping at the first. A fully blank accounted-for row has `BLANK_ROW`, is rejected, and is unquantified. A parseable amount/currency can remain quantified even when another field fails.

Transaction mode key: `(branch_id, transaction_id)`. Compare the six normalized business values, excluding provenance. If all candidates are field-valid and equal, accept the earliest physical imported occurrence and classify the others as `redundant` with `EXACT_DUPLICATE` and its local import reference. Numerically equal exact amounts such as `1.2300` and `1.23` compare equal. If valid payloads conflict, reject every occurrence with `CONFLICTING_TRANSACTION`. If any candidate is invalid, reject the whole readable-key group with `KEY_HAS_INVALID_CANDIDATE` as well as original field issues. A row with an unreadable key is rejected independently; do not guess which transaction it belongs to.

Product mode line key: `(branch_id, transaction_id, line_id)`. Exact duplicate equality compares all normalized product business fields, including the optional normalized label, excluding provenance. Then validate the complete readable transaction group. Any invalid member, conflicting line key, or inconsistent order date/customer/currency rejects **every occurrence in that transaction**, with `TRANSACTION_HAS_INVALID_LINE`, `CONFLICTING_LINE`, or `INCONSISTENT_TRANSACTION` as appropriate. No redundant outcome is assigned within a rejected transaction. Otherwise accept one representative per line key and mark remaining exact occurrences redundant. Different line IDs are not duplicates even when their other values agree. A missing line ID with a readable transaction key invalidates the transaction; an unreadable transaction key cannot invalidate an unidentified group.

Representative selection uses local physical import position, never the optional upstream provenance. Reordering an input may change its duplicate representative, but not accepted business metrics for a valid equal group. There is no latest-file, majority, or correction winner and no cross-run state.

### 5.2 Local quality contract

Every imported data occurrence has exactly one terminal status: `accepted`, `rejected`, or `redundant`. In product mode these counts are **line occurrences**, not transaction counts.

`local_input_count = local_accepted_count + local_rejected_count + local_redundant_count`

Every nonaccepted record has structured issues `{code, field, message}`; redundancy also has a representative reference. Warnings are separate from rejection issues. Local core reason codes include `MISSING_REQUIRED_VALUE`, `ID_NOT_TEXT`, `INVALID_DATE`, `INVALID_AMOUNT`, `UNSUPPORTED_CURRENCY`, `EXCESS_FRACTIONAL_PRECISION`, `AMOUNT_TOO_LARGE`, `INVALID_CELL_TYPE`, `FORMULA_NOT_ALLOWED`, `EXCEL_CELL_ERROR`, `INVALID_QUANTITY`, `NEGATIVE_UNIT_PRICE`, `LINE_AMOUNT_MISMATCH`, plus the group/provenance codes above. Issues may overlap; summed reason counts need not equal rejected record counts.

A local row has a quantified amount only when its `amount` parses and converts exactly to minor units under a known currency precision. Other errors do not erase that amount. Invalid/missing currency or amount, unrepresentable precision, and size overflow make it unquantified. In product mode an invalid unit price/quantity or line mismatch rejects the row/order but does not make an independently valid `amount` unquantified. Track unit-price issues separately; do not count them as missing revenue amounts.

For each known currency and each terminal status record count, sum known amounts and count unquantified rows. Independently check:

`input_known_amount_minor = accepted_amount_minor + rejected_known_amount_minor + redundant_known_amount_minor`

This is an accounting of input occurrences, not a claim that duplicate/excluded amounts are revenue. Unknown currency rows appear in a separate unknown-currency count, not an invented currency bucket with a monetary sum. `local_monetary_complete` means only `local_unquantified_count == 0`. Accepted rows must all have quantified amounts. Never impute zero for unquantified values.

### 5.3 Scope, status, and provenance reporting

Always expose local quality and upstream quality in separate named sections. Do not add the upstream original input count to the local master row count or count upstream exclusions again as local rejections.

Without `--bundle`, upstream input/accepted/rejected/redundant counts, upstream monetary completeness, and upstream acceptance completeness are **unknown / not supplied**, not zero or true. A file called `master.csv` is not proof of a verified SheetMerge run. Missing optional provenance is counted explicitly but is not evidence of missing sales.

With a verified bundle, report upstream counters and financial coverage from section 7. `upstream_acceptance_complete` is a SalesReport Auto derived label meaning `upstream counts.rejected == 0`, not a SheetMerge metadata field. Redundant occurrences remain visible separately. Verification does not establish completeness against the real business population.

`review_required` is true when any local rejection or unquantified amount exists, no local accepted records exist, a provenance/product-label warning exists, a chart cannot be shown safely, or verified upstream data contains rejections or incomplete monetary coverage. Redundancy alone does not require review. Merely omitting a bundle or all optional provenance does not require review; show their unknown/absent scope prominently even on a completed report.

The report-level `review_reasons` list contains each applicable code once, in this fixed order: `NO_ACCEPTED_RECORDS`, `LOCAL_REJECTIONS`, `LOCAL_UNQUANTIFIED_AMOUNTS`, `UPSTREAM_REJECTIONS`, `UPSTREAM_UNQUANTIFIED_AMOUNTS`, `INVALID_PROVENANCE`, `INCOMPLETE_PROVENANCE`, `PRODUCT_LABEL_CONFLICT`, `CHART_PRECISION_LIMIT`. Record-level issue arrays remain separate. This list is empty exactly when a published report has status `completed`.

Successful publication statuses are `completed` (exit 0) and `review_required` (exit 1). Fatal errors exit 2. A review workbook is a usable report of accepted records, never an unconditional all-source-sales total.

## 6. Exact financial and business metrics

### 6.1 Monetary representation and rounding

Let `p` be the resolved currency precision. Parse a valid monetary string as Decimal, then convert it to integer minor units by inspecting its decimal coefficient/exponent, without using the active Decimal context for multiplication. Conversion is valid only when division needed to remove extra fractional digits has zero remainder. Thus `1.2300` at precision 2 is valid; `1.231` is rejected, never rounded. For nonzero amounts and unit prices, cap absolute minor-unit length at 38 decimal digits, matching SheetMerge's row limit. Aggregate sums use unbounded Python integers and can exceed that length. Bound raw monetary tokens as in section 3.2.

All sums and product equality checks use Python integers. Decimal is allowed for exact parsing and explicitly controlled presentation, never a float intermediary. Arithmetic remains correct even if the process Decimal context has precision 2. Never use NumPy fixed-width integers for potentially large monetary sums.

Net revenue is the sum of signed accepted amounts. The term is an application label for this signed sum; it makes no gross-versus-net, accounting revenue-recognition, tax, fee, discount, or settlement claim. Negative values reduce it; zero contributes zero while retaining a transaction/line. Never infer a refund link or classify a negative as a particular accounting event.

Sums display exactly `p` fractional digits. Monetary averages round **once at the final display boundary** to `p` digits using `ROUND_HALF_UP` (ties away from zero). Define rounding on an integer ratio: for numerator `n` in minor units and positive denominator `d`, divide `abs(n)` by `d`; increment the quotient iff twice the remainder is at least `d`; restore the sign. This avoids context-dependent Decimal division. Retain the exact numerator/denominator in the logical metric result for verification. A zero denominator displays `N/A`, not 0.00, blank chart zero, or an exception. There is no rounding of row values into acceptance, and no averaging already-rounded averages.

### 6.2 Common metrics

All metric groups are computed from accepted representatives only, within **one currency**. A transaction is identified by `(branch_id, transaction_id)`. In product mode its total is the sum of its accepted lines; group validation ensures it has one date, customer, and currency.

| Metric | Exact definition |
| --- | --- |
| `accepted_record_count` | Number of accepted transaction rows in transaction mode or accepted line rows in product mode |
| `transaction_count` | Number of distinct accepted transaction keys, including zero-total and negative-total transactions |
| `unique_customer_count` | Number of distinct `(branch_id, customer_id)` pairs represented by accepted transactions in the group; a conservative branch-scoped identity rule, not a global person count |
| `net_revenue` | Sum of accepted transaction amounts, or accepted line amounts in product mode; these give the same product-mode transaction total |
| `average_transaction_value` | Group net revenue divided by group transaction count, including zero and negative transactions in the denominator; final rounding as above |

Summary groups by currency. Monthly Sales groups by `(currency, transaction_date year-month)` using calendar months and displays `YYYY-MM`; include only months with accepted transactions, with no manufactured zero months. Branches groups by `(currency, branch_id)`. Recompute distinct counts and averages in each group. Monthly/branch revenue and transaction counts partition Summary; distinct customers generally do not add across months or branches. A customer's repeated lines count once within each applicable group.

### 6.3 Product metrics

Products groups accepted lines by `(currency, product_id)` across branches. Never combine different quantity units under a product ID; schema users are responsible for a consistent whole-unit definition.

| Metric | Exact definition |
| --- | --- |
| `accepted_line_count` | Number of accepted representative lines for the product |
| `transaction_count` | Distinct accepted transaction keys containing that product, including zero/negative lines |
| `unique_customer_count` | Distinct branch/customer pairs on those transactions |
| `positive_quantity` | Sum of quantities strictly greater than zero |
| `negative_quantity_magnitude` | Sum of absolute quantities strictly less than zero |
| `net_quantity` | Signed sum of quantities, equal to positive quantity minus negative magnitude |
| `net_revenue` | Signed sum of this product's accepted line amounts |
| `minimum_unit_price`, `maximum_unit_price` | Minimum/maximum recorded accepted unit price, including zero/negative-quantity lines; no averaging |
| `positive_quantity_weighted_unit_price` | Sum of amounts on strictly positive-quantity lines divided by positive quantity, in minor units per unit; final rounding to `p`; `N/A` if positive quantity is zero |
| `revenue_rank` | Sequential position starting at 1 within the currency, sorted by net revenue descending then exact product ID ascending; ties get distinct deterministic ranks |

Product revenue and line counts add to Summary. Product transaction/customer counts can overlap and must not be summed to obtain Summary. No generic mean of price cells, net-revenue/net-quantity price, inferred profit, revenue share percentage, or cross-currency ranking is included in MVP. Negative and zero product revenue stay in the ranking.

## 7. Optional SheetMerge bundle adapter

### 7.1 Selection and verification

`--bundle DIRECTORY` is supported only with `--mode transactions`. Require the seven files named in section 2. The selected input must be named `master.csv` or `master.xlsx`; it may be a byte-identical copy outside the bundle. Hash the input snapshot and match the corresponding `manifest.output_sha256` entry. Do not automatically consume both masters or merge exclusions into sales.

Parse JSON strictly: reject duplicate object keys, nonstandard NaN/Infinity, wrong types (including booleans used as integers), missing required fields, unsupported `manifest_version`, malformed digest strings, and negative counts. Allow extra object fields for additive compatibility, but never use unrecognized fields to establish completeness. Record `tool_version`; do not promise compatibility with a new manifest version merely because a version string looks similar.

Require `tool == "sheetmerge"`, `manifest_version == 1`, and `business_key == ["branch_id", "transaction_id"]`. Require the key set of `output_sha256` to be exactly the six non-manifest output basenames in section 2; each digest is 64-character lowercase hexadecimal SHA-256 text. Reject extra keys, including path-like names. Hash those six bundle members and compare. Do not follow symlinks/junctions, interpret manifest keys as arbitrary paths, or open original sources from `inputs`/provenance. The manifest has no self-hash or signature; these checks establish internal consistency with the supplied manifest, not authenticity.

The mandatory adapter parses `manifest.json`, `reconciliation.json`, and `exclusions.csv`. `reconciliation.csv` and `audit.jsonl` are hash-checked companion artifacts, not metric inputs; the alternative master is hash-checked but need not be decoded. Their actual schemas are documented below to avoid invented field names. Deep audit replay and reconstruction of original source files are outside MVP.

### 7.2 Actual manifest and reconciliation schemas

Required manifest fields consumed/retained are:

- `tool`, `tool_version`, `manifest_version`, `business_key` as above.
- `configuration.currency_precision`: nonempty object mapping uppercase currency codes to integer precision 0-9. `configuration.profiles` is an array and `configuration.limits` an object; retain their JSON as source configuration without executing/reapplying source mappings.
- `configuration_sha256`: SHA-256 text; retain as a declared configuration fingerprint, not a recomputed proof of configuration equivalence.
- `inputs`: array of objects with `source_file`, `sha256`, `profile`, all text; source filenames must be unique. Preserve names/hashes, without accessing the source files.
- `policies`: object with actual keys `duplicates`, `conflicts`, `invalid_candidates`, `blank_rows`, whose values are explanatory strings. Retain them; they are not executable settings.
- `reconciliation`: the same object schema as `reconciliation.json`; require semantic equality of the parsed objects.
- `output_sha256`: the basename-to-digest map described above. There is no persisted `status`, `output_directory`, or `monetary_complete` at manifest root; coverage lives under `reconciliation`.

The reconciliation root has these exact fields:

| Field | Type and meaning |
| --- | --- |
| `counts` | Object with nonnegative integer `input`, `accepted`, `rejected`, `redundant` |
| `count_equation` | String `input = accepted + rejected + redundant` |
| `count_balanced`, `known_amounts_balanced` | Booleans that must be true for a supported published bundle |
| `monetary_complete` | Boolean equal to whether `unquantified_count` is zero; not whether all rows were accepted |
| `unknown_currency_count`, `unquantified_count` | Nonnegative integers, each at most input; unknown currency count at most unquantified count |
| `currencies` | Array with one unique entry for every configured currency, including currencies with zero rows |

Each currency entry has `currency` (text), `precision` (integer), `known_amounts_balanced` (boolean), `monetary_complete` (boolean), and four objects keyed by `input`, `accepted`, `rejected`, `redundant`: `counts` (integers), `amounts` (fixed-precision signed decimal strings), `minor_units` (signed integer strings), `unquantified_counts` (integers).

Validate global and per-currency count partitions, known-amount partitions, and unquantified partitions. Check nonnegative counts and that unquantified counts do not exceed corresponding counts. Check that every amount string exactly represents its matching minor units at `precision`, and precision matches configuration. Accepted unquantified counts must be zero. Validate each currency's monetary-complete flag against its input unquantified count; validate global completeness likewise. Sum currency input counts plus unknown-currency count to global input; sum currency input unquantified counts plus unknown-currency count to global unquantified count. Sum per-currency accepted counts to global accepted. Other per-status currency counts cannot exceed global corresponding counts.

The CSV companion's exact column groups, in order, are `currency`, `{input,accepted,rejected,redundant}_count`, `{input,accepted,rejected,redundant}_known_amount`, `{input,accepted,rejected,redundant}_unquantified_count`, `known_amounts_balanced`, `monetary_complete`. All CSV cells are text; boolean spelling is lowercase `true`/`false`. JSON uses `amounts` and `minor_units`, not CSV's `*_known_amount` field names.

### 7.3 Actual exclusions and audit schemas

`exclusions.csv` headers are, in order: `branch_id, transaction_id, transaction_date, customer_id, amount, currency, source_file, source_sheet, source_row, status, reason_codes, reasons, duplicate_of_file, duplicate_of_sheet, duplicate_of_row, raw_values`.

Parse it as strict CSV text. `status` must be `rejected` or `redundant`. `reason_codes` is a nonempty JSON array of code strings; `reasons` is a nonempty JSON array of objects with `code`, `field`, `message`; require corresponding code lists to agree. Unknown nonempty reason codes remain visible, not discarded. `raw_values` is a JSON object keyed by original source headers. Invalid canonical conversions can be empty strings; parseable amounts may remain present on rejected records. Do not run accepted-sales validation on these rows or treat an empty/invalid amount as zero. Preserve the original canonical/raw data and source reference.

Redundant rows supply `duplicate_of_file`, `duplicate_of_sheet`, `duplicate_of_row`; rejected rows leave these empty. Source/duplicate row numbers are positive decimal-digit text; sheets can be empty for CSV. Require each redundant reference to identify an accepted selected-master provenance reference, and require each source file to occur in `manifest.inputs`. Enforce unique provenance across selected master and exclusions, which together cover upstream input occurrences. Counts of each exclusion status, globally and by known currency, must match the corresponding reconciliation counters. Global unquantified/known rejected ledgers remain upstream declarations checked for internal reconciliation, not independently reconstructed from the excluded raw cells.

The actual audit JSONL object fields are `source`, `status`, `values`, `amount_minor`, `issues`, `duplicate_of`, `raw_values`, `raw_cells`. `source` and non-null `duplicate_of` contain `source_file`, `source_sheet`, `source_row` (integer). `status` includes accepted/rejected/redundant; normalized dates and Decimal amounts are strings; `amount_minor` is an integer string or null. `issues` contains code/field/message objects. Each raw-cell object has `value`, `kind`, `numeric_text`, `invalid_excel_date`, `original_text`; raw numeric audit representations are not authoritative operands. The adapter must not assume these fields exist in master CSV/XLSX.

### 7.4 Selected-master agreement and quality propagation

Require the selected master to have exactly the nine SheetMerge columns and complete valid provenance. Require every row to survive local business validation with no rejection or redundancy, and to reference a declared input source. Local accepted count must equal upstream accepted count; per-currency accepted counts and exact sums must equal reconciliation accepted ledgers. Missing rows, local conflicts, invalid values, or mismatched totals are fatal `BUNDLE_MISMATCH`, not a reason to silently downgrade to standalone import. Hash/metadata failures are fatal `INVALID_BUNDLE`.

Publish local and upstream quality separately. In particular, for the known sample, local counts are 16 input / 16 accepted / 0 rejected / 0 redundant; upstream counts are 33 / 16 / 14 / 3. Propagate upstream review because rejected=14 and `monetary_complete=false`. Do not claim that all 33 records contributed to revenue or that five missing/unrepresentable amounts were zero.

## 8. Excel deliverable and deterministic presentation

### 8.1 Workbook sheets

The one output file is an `.xlsx` workbook. Visible sheet order is exactly the table below, with Products omitted in transaction mode. Use styled headers, frozen table headers, filters on data tables, readable bounded widths, and currency/units in titles. No macros, external links, refresh requirement, or Excel formulas are necessary for business metrics.

The mandatory scope statement is: "Metrics include accepted records only. Precise totals do not imply that all source records were accepted." Without a bundle add: "Upstream acceptance and monetary completeness were not supplied."

| Sheet | Mandatory contents |
| --- | --- |
| `Summary` | Report status and review reasons; mode, application/spec version; input basename/SHA-256, chosen sheet, effective precision table; provenance coverage; local counts and monetary coverage; upstream status/counters or explicit unknowns. Per observed accepted currency: `currency`, `precision`, `accepted_record_count`, `transaction_count`, `unique_customer_count`, `net_revenue`, `average_transaction_value`. Display the accepted-record scope statement beside metrics. |
| `Monthly Sales` | `currency`, `month`, `accepted_record_count`, `transaction_count`, `unique_customer_count`, `net_revenue`, `average_transaction_value`; one row per observed accepted currency/month |
| `Branches` | `currency`, `branch_id`, `accepted_record_count`, `transaction_count`, `unique_customer_count`, `net_revenue`, `average_transaction_value` |
| `Products` | Product mode only: `currency`, `revenue_rank`, `product_id`, `product_name`, and all product metrics from section 6.3 (do not duplicate the rank column); no placeholder product sheet in transaction mode |
| `Data Quality` | Separate local/upstream summary tables with input/accepted/rejected/redundant counts, unknown-currency and unquantified counts, coverage flags, count/known-amount balance checks; upstream unknowns if absent. Per-currency local and supplied upstream ledgers; provenance coverage counts; detailed excluded/warned occurrences and review reasons; declared source policies/version and verification scope if supplied. |
| `Raw Data` | Every local input occurrence, including rejected/redundant rows, preserving all original input column names/order as literal values; append `report_import_file`, `report_import_sheet`, `report_import_row`, `report_status`, `report_reason_codes`, `report_issues`, `report_warnings`, `report_duplicate_of`, `report_amount_minor`, `report_currency_precision`, `report_normalized_date`, and `report_raw_cell_kinds`. Product mode additionally appends `report_quantity`, `report_unit_price_minor`. Missing derived values stay empty, not zero. |

Data Quality detail columns are `scope`, `import_file`, `import_sheet`, `import_row`, `source_file`, `source_sheet`, `source_row`, `status`, `reason_codes`, `issues`, `warnings`, `duplicate_of`, `raw_values`. Use one detail row per local excluded/warned occurrence and one per upstream exclusion; absent local references on upstream rows remain empty. Serialize structured cells as JSON. Include complete original local values in Raw Data and upstream excluded `raw_values` in Data Quality. Keep upstream exclusions separate from local Raw Data: the input master never contained them.

All identifiers, source strings, raw values, dates, and **authoritative monetary values** are explicitly written as literal Excel text, including strings beginning with `=`, `+`, `-`, or `@`. Do not prepend apostrophes that change identifier content or allow openpyxl to create formula cells. Counts/whole quantities within Excel's exact integer range can be numeric; larger integers and minor-unit strings remain text. Monetary averages are displayed rounded values with the rule documented in the workbook. `N/A` is literal text.

Do not rely on Excel SUM over text amounts. The application's verified values are authoritative. Reject output values exceeding Excel's 32,767 UTF-16-unit cell length, unsupported XML characters, or 1,048,576 worksheet rows (including headings/sections). Fail publication instead of truncating audit information. Synthetic header-only/all-rejected reports retain the applicable empty table headers.

### 8.2 Charts

For each currency with accepted data, create a line chart on Monthly Sales showing net revenue by observed month and a column chart on Branches showing net revenue by branch. Product mode adds a horizontal bar chart on Products of up to ten products in revenue-rank order. A single observation is valid. Retain negative/zero values; use no pie chart and no stacked mixed-currency total. Every title includes its currency; one chart never includes another currency as a series or category.

Chart data uses clearly labelled helper columns on the corresponding sheet: integer minor units and an axis caption such as `Net revenue (USD minor units; 100 = 1 USD)`. These helpers are generated from exact aggregates, not recomputed with Excel or floats. Allow a currency chart only if every plotted absolute minor-unit integer is at most 999,999,999,999,999 (15 digits). Otherwise omit that chart, record `CHART_PRECISION_LIMIT`, and require review; keep exact table totals. Chart rendering inside Excel is presentation, not an additional exact financial calculation. No conversion to Python float is required even for chart source values. Skip charts entirely when no accepted data exists.

### 8.3 Ordering and repeatability

- Currency rows/charts: ascending exact currency code. Months: chronological within currency. Branches: currency then exact branch ID ascending. Products: currency then the rank order defined in section 6.3.
- Raw Data: original physical import order. Quality details: local before upstream, then source/import filename, sheet, numeric row, with a fixed tie-breaker of original occurrence index. Issues/warnings within a row: field-schema order then code. Ledger columns always use input, accepted, rejected, redundant order.
- Treat all ID sorts as Unicode lexical, not numeric/natural or locale-dependent. Sorting does not trim or fold case. Group keys are exact preserved strings.
- Fixed inputs/options produce identical logical tables, statuses, sheet/chart order, and rounded display values. Byte-identical ZIP files and timestamps are not an MVP promise. Do not put current time into calculations or ordering.

## 9. CLI, file safety, and publication

The installed command will be `salesreport-auto`; `python -m salesreport_auto` exposes the same interface. There are no subcommands in MVP.

```text
salesreport-auto --help
salesreport-auto --version
salesreport-auto --mode {transactions,products} --input FILE --output REPORT.xlsx
                [--bundle DIRECTORY] [--currency-precision CODE=N ...]
```

`--help` documents schemas, worksheet selection, precision defaults, review statuses, and refusal to overwrite. `--version` prints application version without requiring other arguments. `--mode`, `--input`, and `--output` are required for a report. `--output` is a file path, not SheetMerge's new-directory argument. All optional and required arguments may appear in any order; unknown options are errors. There is no `--force`, automatic schema detection, or silent fallback after a bad bundle.

Examples (illustrative commands, not actions performed by this specification):

```text
salesreport-auto --mode transactions --input sales.csv --output reports/sales.xlsx
salesreport-auto --mode transactions --input handoff/master.xlsx --bundle handoff --output reports/merged.xlsx
salesreport-auto --mode products --input lines.xlsx --output reports/products.xlsx --currency-precision EUR=2
```

Input, bundle artifacts, and source repositories are opened read-only. Validate resolved paths and reject symlink/junction inputs/bundle members or output aliases. The output must have `.xlsx` extension and an existing writable parent directory. Reject an existing target of any kind, including a dangling link; do not change it. Resolve identity to prevent output/input collision (including hard-link aliases); refuse any output inside a supplied bundle directory. The output may be a new distinct file beside a standalone input. No deletion/moving of inputs or previous reports is permitted.

Read each consumed file into a stable byte snapshot or stream with identity/hash checks sufficient to detect changes during the run; use the bytes validated for parsing/hashing, not a later reopened different version. A changed input/bundle member fails the run. Do not log raw customer records on stdout/stderr.

Build the report in a uniquely named temporary file in the output parent. Reopen it with openpyxl to verify sheet names/order, metric values, row counts, literal text cells, and chart partition/source references against in-memory results. Recheck local conservation and aggregate partitions. Only then publish using a no-clobber operation that fails if the target has appeared concurrently; an ordinary replacing rename is insufficient. Clean up only this invocation's temporary file on failure. Do not leave a final partial workbook.

Successful publication prints one JSON object on stdout with fields `status` (`completed` or `review_required`), `mode`, `output`, `local_counts` (input/accepted/rejected/redundant), `local_unquantified_count`, `local_monetary_complete`, `upstream` (null without bundle, otherwise counts/unquantified_count/monetary_complete), `review_reasons` (ordered code list), and `currencies` (code/precision/net_revenue string/transaction_count). These are SalesReport Auto fields, not asserted SheetMerge manifest fields. Exit codes follow section 5.3. Fatal errors print `salesreport-auto: error: ...` on stderr, with file/row/field context when available, and exit 2 without a success JSON object. Argparse help/version exit 0; argument errors exit 2.

## 10. Technical design and feasible MVP boundary

Use a small synchronous in-memory pipeline: parse arguments and protect paths; load snapshots and optional bundle declarations; parse typed raw records; validate fields; resolve transaction/line groups; reconcile local classifications; verify bundle agreement; compute integer metrics; create and read back workbook; publish. Keep parsing, classification, aggregation, and presentation as separately testable responsibilities. No detailed module layout or code is prescribed at this stage.

Pandas may organize object/string columns and group indexes; Python integer accumulators/sets compute monetary sums and distinct keys. Openpyxl provides workbook structure/charts. Avoid new runtime dependencies, Excel/COM automation, databases, a server/UI, and invoking SheetMerge in normal operation. Preserve explicit raw/local/upstream identity rather than mutating source rows into accepted data.

Mandatory single-day scope is deliberately narrow: one file, fixed schemas/text XLSX inputs, declared precision, whole-unit product lines with exact amount-price agreement, basic fixed Excel tables/charts, record-level quality, the version-1 bundle adapter, safe publication, and targeted pytest/Ruff checks. No source-dialect detection, numerical XLSX money decoder, original-source audit replay, arbitrary field mapping, pivot tables, or configurable report designer is needed.

Future enhancements, none required for MVP:

- Additional input profiles, native XLSX numeric/date cells with exact stored-value decoding, locale-specific CSV formats, or mapping configuration.
- Fractional quantities, sub-minor unit prices, discounts, tax/gross/net models, explicit refund links, product units/catalog dimensions, and agreed price rounding policies through a new schema version.
- Cross-branch customer identity, exchange rates/base-currency reporting, consolidated currency charts, margins/costs, forecasts, or percentages.
- Multi-file/incremental imports with persisted idempotency, correction workflows, and order-line completeness manifests.
- Parsing audit.jsonl, rechecking rejected monetary ledgers against original raw cell metadata, verifying source/configuration fingerprints independently, digital signatures, or hardened hostile-file processing.
- Interactive dashboards, PDF/email delivery, branding/templates, richer charts, byte-reproducible workbooks, or performance beyond bounded local batches.

## 11. Automated acceptance plan (to implement later)

Default tests must create their own synthetic files under pytest temporary directories and run without the sibling repository, network, or preinstalled SheetMerge. Parametrize applicable cases across CSV/text-XLSX. No tests or fixtures are implemented by this specification.

### 11.1 Standalone tests and concrete expected outcomes

| ID | Input or action | Required assertions |
| --- | --- | --- |
| S01 | Minimal six-column transaction CSV and equivalent text-XLSX `Master` | Same accepted metrics; no bundle dependency; provenance absent/unknown; five sheets in defined transaction order; no Products |
| S02 | IDs `001`, `000001`, `NA`, long digit text, Unicode, and literal `=1+1`; optional blank `source_sheet` | IDs unchanged, no NA inference or formula cells; complete source file/row plus blank sheet is valid provenance |
| S03 | USD transactions at branch A: `10.00` and `-2.00` for customer C1, `0.00` for C2, all January; branch B `3.01` for C1 in February | Four transactions, three branch/customer identities, USD revenue `11.01`, ATV `2.75`; January count 3/revenue `8.00`/customers 2/ATV `2.67`; February count 1/revenue `3.01`; branch totals agree; negatives/zeros retained |
| S04 | Two accepted USD amounts `0.00`, `0.01`; separately `0.00`, `-0.01` | Exact half-up ATVs `0.01` and `-0.01`; no rounding of row inputs |
| S05 | USD `0.10` repeated on ten different keys; amounts with 30+ digits; context precision deliberately set to 2 | Sum `1.00` and exact large sums independent of Decimal context; no float/64-bit overflow; oversized chart omitted with review but exact table retained |
| S06 | Monetary `1.2300`, `1.231`, `NaN`, `Infinity`, `1e2`, `1,23`, whitespace, missing amount, 39-digit nonzero minor units | First accepted as `1.23`; others rejected with specific reasons, unquantified amounts not zero; size/precision boundaries tested at exactly allowed limits |
| S07 | Real leap date `2024-02-29`; invalid `2023-02-29`, `2024-2-29`, timestamps, and native XLSX date | Only ISO leap date accepted; invalid-date rows with valid money remain quantified; no timezone/month inference |
| S08 | USD/PLN/JPY rows and unknown `EUR`; rerun with `--currency-precision EUR=2` | Currency totals/charts remain separate; EUR initially rejected/unquantified then accepted; JPY fractional nonzero precision rejected; duplicate/invalid precision options fail |
| S09 | Three identical normalized same-key transactions; then same-key amount conflict; then same-key invalid date | First case accepted=1/redundant=2 with earliest physical representative; conflict/invalid cases reject all occurrences; same transaction ID in another branch remains independent |
| S10 | Missing customer/key, partial row, fully blank row, blank CSV record, quoted multiline field | Independent issues retained; one classification per physical record; blank rejected/unquantified; multiline record counted once with correct start line |
| S11 | Zero bytes; malformed CSV/encoding; missing/extra/duplicate headers; corrupt XLSX; multiple ambiguous sheets; data beyond headers; limit breaches | Fatal exit 2 and no final workbook; no silent column/row deletion; `Master` selection with extra sheets is explicit and listed |
| S12 | Valid headers with no records; valid records all rejected | Exit 1; Summary states no accepted records; applicable sheet/table headers remain; no currencies/charts invented; averages absent or `N/A`, never fabricated zero |
| S13 | Missing all provenance; valid complete reference; partial reference; invalid row text | First informational; second preserved; partial/invalid warnings trigger review without rejecting valid sales; local import location separate from upstream reference |
| S14 | Numeric/boolean/formula/error XLSX business cells and optional numeric source row | Business values rejected with cell-type reasons; formula never evaluated; exact positive numeric source row permitted; text values preserved literally in output |
| S15 | Mixed accepted/rejected/redundant rows, including rejected invalid-date rows with money | Count and per-currency known-money equations hold; accepted totals exclude all exclusions; complete monetary coverage can coexist with rejected records and review |
| S16 | Product order T1: P1 line L1 quantity `2`, price `3.00`, amount `6.00`; P2 L2 `1`/`4.00`/`4.00`; order T2 P1 L1 `-1`/`3.00`/`-3.00`; one USD branch/customer | Three accepted lines, two transactions, one customer, revenue `7.00`, ATV `3.50`; P1 revenue `3.00`, net quantity 1, positive 2, negative magnitude 1, weighted price `3.00`; P2 ranks 1 with `4.00`, P1 ranks 2 |
| S17 | Product lines with repeated identical line IDs; distinct line IDs with equal values; conflicting line; invalid quantity/price/date; inconsistent order currency/customer/date | Exact repeats redundant only in a valid transaction; distinct line IDs count separately; invalid/conflicting group quarantines whole readable transaction with no partially counted order; valid other transactions survive |
| S18 | Product zero quantities/prices, negative quantities, amount-price mismatch, fractional quantity, sub-minor price, no positive quantities | Exact valid zeros/reversals accepted; mismatch/fractional quantities/sub-minor prices rejected; weighted price `N/A` with no positive quantity; independent amount quantification still correct |
| S19 | Same product across orders/branches, tied product revenue, differing labels | Exact product grouping per currency; transaction/customer counts not summed across products; tie rank by product ID; deterministic label selection with warning/review |
| S20 | SheetMerge-shaped master in product mode; product schema in transaction mode | Both fail headers clearly; no manufactured product fields or ignored line columns |
| S21 | Existing report, output=input, hard-link alias, dangling link, output inside bundle, publication collision, writer/read-back failure | Existing/source bytes and mtimes unchanged; no overwrite/partial final file; only invocation-owned temporary file cleaned up |
| S22 | Workbook read-back after mixed-currency report | Required sheet names/order; exact text totals/IDs/raw values; all local input rows retained; structured reasons; chart titles/data references each contain only their currency; integer helper limit honored |
| S23 | Repeated runs with fixed inputs/options; reordered valid business rows | Identical logical metrics and specified ordering for fixed input; row permutation preserves financial metrics but can change raw order/duplicate representative; no dependency on locale or run time |
| S24 | CLI help/version/missing or unknown arguments; completed/review/fatal runs | Correct exit codes; help/version need no input; documented JSON types and ordered review reasons; fatal error uses stderr and publishes nothing |

### 11.2 Adapter unit tests independent of SheetMerge

Create small synthetic bundles within SalesReport Auto tests following the **actual** schemas in section 7, not the 33-record fixture expectations. Test valid custom precision (for example `KWD=3`), fixed-precision/minor-unit agreement, unknown original-currency counts, parsed exclusions and representative references, null-free integer counters, and local/upstream count separation.

Tamper independently with selected master bytes, a companion digest, required files, manifest version/tool/key, configuration precision, embedded reconciliation, count or amount partitions, coverage flags, exclusion counts/codes, source inventory, and duplicate references. Every inconsistent bundle must fail with exit 2 and no final report, without fallback. Test JSON duplicate keys/booleans-as-counts, malformed nested JSON in exclusions, and unsafe manifest filenames. A valid upstream conflict-only bundle may have zero accepted rows and `monetary_complete=true`; its empty report still requires review because upstream rejections/no accepted records exist.

A standalone clean input must not acquire upstream counts/completeness from its filename, provenance columns, neighboring files, or remembered fixture expectations. Verify that omitting `--bundle` leaves upstream JSON null and workbook upstream fields unknown, including for a file with the same bytes as a master.

### 11.3 Explicit sibling-fixture integration tests

Mark these as opt-in integration tests, skipped with a clear reason if `../sheetmerge` and its compatible environment are unavailable. They must not be necessary for the default standalone test suite. Generate any SheetMerge outputs in pytest's isolated temporary directory, never in either repository's sample/output folders; read the sibling samples without modifying them. Tests may invoke the sibling's CLI solely as integration setup, not as SalesReport Auto runtime behavior.

Use SheetMerge's `samples/sources.toml` with `samples/input/processable`. Expect SheetMerge exit 1. Test SalesReport Auto with each of the resulting CSV and XLSX masters separately, first with its bundle and then without. The fixture-level assertions are:

| Scope | Expected result |
| --- | --- |
| Original upstream inventory | 11 processable files; 33 occurrences |
| Upstream classifications | accepted=16, rejected=14, redundant=3 |
| Selected-master local import | input=16, accepted=16, rejected=0, redundant=0, local unquantified=0 |
| Accepted USD | `141.45`; 7 accepted transactions |
| Accepted PLN | `220.35`; 5 accepted transactions |
| Accepted JPY | `1290`; 4 accepted transactions |
| Upstream monetary coverage | unquantified=5, unknown currency=0, `monetary_complete=false` |
| Bundle report status | review required, exit 1; original rejections/redundancy shown separately |
| Same master standalone | Same accepted metrics; upstream counts/coverage unknown; complete optional provenance retained; exit 0 absent other warnings |
| Workbook structure | Transactions sheets only; no Products; separate USD/PLN/JPY charts; all metrics confined to accepted rows |

Check upstream known-money ledgers exactly: USD input `356.55` = accepted `141.45` + rejected `115.00` + redundant `100.10`; PLN `583.35` = `220.35` + `163.00` + `200.00`; JPY `7389` = `1290` + `5099` + `1000`. Check reason occurrences from `samples/expected.json`: `CONFLICTING_TRANSACTION=6`, `EXACT_DUPLICATE=3`, `EXCESS_FRACTIONAL_PRECISION=2`, `KEY_HAS_INVALID_CANDIDATE=2`, `INVALID_DATE=3`, `MISSING_REQUIRED_VALUE=3`. Their overlap is intentional.

Separately use SheetMerge's complete `samples/input` tree, which includes the missing-column workbook. Expect SheetMerge exit 2, missing `account_id` and `currency_code`, and no completed bundle. Do not assert that this structurally failed 36-physical-record input has the processable 33-record result. SalesReport Auto must not invent a report from an absent bundle/master. Product tests use SalesReport Auto's own line-item test data, never claims that these SheetMerge fixtures contain products.

The final implementation verification gate is the default pytest suite, opted-in fixture integration when available, Ruff, and workbook read-back assertions. Broader performance/hardening work is future scope.

## 12. Meaning and limits of a successful report

A successfully published report guarantees that its accepted records meet the declared mode/precision contract, are counted once under the documented key policy, use exact monetary sums, are partitioned by currency, and are accompanied by visible quality/provenance scope. Local count and known-money conservation checks passed. The saved workbook matches the computed logical report. With `--bundle`, the documented manifest/hash/schema/ledger agreement checks passed and supplied upstream quality was propagated.

It does not guarantee that every input occurrence was accepted, that no original sales are missing, that an upstream file/bundle represents the entire business, or that source-provided IDs, dates, amounts, labels, quantities, and provenance are true. Exact accepted totals do not imply acceptance completeness. No bundle means upstream completeness is unknown; a complete upstream monetary ledger can still contain rejected transactions. Hash consistency is not authentication or an external financial audit.

It does not infer taxes, gross/net semantics, refund links, accounting periods beyond calendar-month grouping, currency exchange rates, missing customer identities, product fields absent from the input, or unprovided product lines. Product-mode all-or-nothing validation protects the lines supplied under a readable transaction key; it cannot detect a line that was never supplied. Charts and displayed rounded averages do not replace the exact underlying integer aggregates and documented denominators.
