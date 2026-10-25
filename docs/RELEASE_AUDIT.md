# SalesReport Auto 0.1.0 release audit

Audit date: 2026-10-25. Scope: local source, specification 1.0, tests, documentation,
distributions and installed-package smoke tests. This is a portfolio release review,
not production certification or an independent financial audit.

Publication documentation note (2026-10-25): the README was subsequently revised
for portfolio readers, including reproducible demos, test evidence, release status
and the unselected license. Application behavior, tests, package metadata and the
specification were not changed by that documentation pass. Distribution results
below describe the audit's earlier builds; rebuild local archives to include the
updated documentation before distributing them.
The documentation pass reran the default suite (521 passed, 4 skipped, 0 failed),
opt-in sibling suite (3 passed), Ruff and pip check successfully. All eight README
links resolved locally; seven PowerShell blocks parsed without errors. Six report
examples, both inline CSV schemas and the read-only Python example were exercised
using isolated outputs and synthetic replacements for placeholder input files.

## Recommendation

Technical release candidate, subject to the verification limits and documentation
clarifications below. No unresolved financial-calculation or no-clobber defect was
found after the focused fixes. Passing tests are evidence for their scenarios,
not proof of correctness for every input or of original business completeness.
Do not describe this software as production-certified or financially audited.

No commit, push, PyPI publication or SheetMerge repository edit was performed.
The repository's application/docs/tests were already untracked at audit start;
the audit leaves them untracked. There is no committed release snapshot to cite.

## Verification results

| Check | Result |
| --- | --- |
| Initial default pytest baseline | 492 passed, 4 skipped, 0 failed |
| Final default pytest | 521 passed, 4 skipped, 0 failed |
| Opt-in sibling SheetMerge integration | 3 passed, 0 skipped, 0 failed |
| Ruff (`python -m ruff check src tests`) | Passed |
| Project virtual-environment `pip check` | Passed; no broken requirements |
| Wheel and source distribution | Both built successfully; metadata/content checks passed |
| Tests from extracted source archive | 521 passed, 4 skipped, 0 failed |
| Installed-wheel API/module/console smoke tests | Passed; both modes and both entry forms, help/version and collision checks |

Default skips: one real symlink test cannot create symlinks with the current
Windows permissions; three SheetMerge cases require explicit opt-in. Simulated
link/junction guards still run by default. The opt-in cases ran successfully in
the available compatible sibling environment, including CSV and XLSX master
imports and the deliberate structurally invalid source tree. Sample bytes and
modification times were checked unchanged; sibling subprocess bytecode writes
were disabled. All generated bundles/reports went into isolated pytest folders.

The real processable fixture produced 33 original occurrences: 16 accepted,
14 rejected, 3 redundant; 5 unquantified amounts, incomplete upstream monetary
coverage, and review required. Accepted revenue: USD `141.45`, PLN `220.35`,
JPY `1290`; transaction counts 7, 5 and 4 respectively. Read-back checked the
five sheets, upstream ledgers, 17 exclusions, 16 local master rows and separated
chart references. The same masters standalone retain accepted metrics and expose
unknown upstream scope, rather than inheriting fixture completeness.

Environment: Windows, Python 3.12.15, openpyxl 3.1.5, pandas 3.1.0,
pytest 9.1.1, Ruff 0.17.2. Build frontend 1.6.1, setuptools 84.2.1 and wheel
0.48.0 were installed in the project virtual environment for packaging checks.
The initial sandboxed download attempt failed because network access was denied;
the authorized tool installation then succeeded. Normal reporting/default tests
require neither downloads nor a sibling installation.

Reproduction from the repository root (fresh basetemp paths avoid restricted
system-temp access and pytest deleting an earlier run):

```powershell
$run = Join-Path (Get-Location) ('.test-tmp\audit-' + [guid]::NewGuid().ToString('N'))
.venv\Scripts\python.exe -m pytest -q --basetemp=$run --tb=short
$integration = Join-Path (Get-Location) ('.test-tmp\integration-' + [guid]::NewGuid().ToString('N'))
.venv\Scripts\python.exe -m pytest -q tests/test_sheetmerge_integration.py --sheetmerge-integration --basetemp=$integration
.venv\Scripts\python.exe -m ruff check src tests
.venv\Scripts\python.exe -m pip check
.venv\Scripts\python.exe -m build --no-isolation
```

`--no-isolation` was used with the explicitly installed build backend; a new
development installation can instead use `python -m build` and its isolated
build environment, which may download dependencies.

## Mandatory requirements to evidence

Status **I** means the implementation was inspected and implements the stated
behavior; **P** would mean incomplete implementation; **U** means absent.
An I is not a claim of exhaustive automated coverage. Evidence paths are relative
to this repository. Tests named below are representative direct assertions;
parameterized cases also cover CSV/text-XLSX where applicable.

| Contract | Status | Implementation evidence | Direct automated evidence |
| --- | --- | --- | --- |
| 3.1 single bounded CSV/XLSX snapshot; strict decoding, headers, physical rows | I | `readers.py`, `config.py` | `test_readers_provenance.py`: `test_empty_malformed_encoding_and_width_errors`, `test_s11_bad_headers_fail_structurally`, `test_record_byte_and_archive_limits`; `test_transactions.py`: S10 |
| 3.1 exact worksheet selection, sparse spans, ignored worksheets, empty cells | I | `readers.py`; `report_model.py:_quality` | Reader selection/stale-dimensions/sparse-span tests; `test_release_audit.py`: empty text/header cases and `test_ignored_worksheets_are_disclosed_in_data_quality` |
| 3.1 unsupported/encrypted/unsafe/duplicate ZIP entries; no evaluation | I | `readers.py:_check_archive`, `_read_xlsx` | Reader compression/corrupt-workbook tests; release audit malformed-package cases; S14 formula/error cases |
| 3.2 exact text/Unicode/NA/IDs; date and type rules | I | `models.py`, `readers.py`, `validation.py` | S02/S07/S14; `test_formula_like_identifiers_and_raw_text_remain_exact`; product literal-text/read-back cases |
| 3.3 defaults, explicit precision, complete authoritative bundle mapping | I | `config.py`, `cli.py`, `bundle.py` | `test_money_config.py`; S08; CLI invalid/repeated declarations; bundle custom/default precision tests |
| 4.1 six transaction columns, optional provenance and coverage | I | `models.py`, `validation.py:_provenance` | S01/S13/S14; reordered/optional headers, numeric source-row and provenance-type cases |
| 4.2 explicit product-line schema, quantities, exact price equality, labels | I | `product_validation.py`, `models.py` | `test_products.py`: S16/S18/S19/S20, required headers, price/quantity boundaries and label-length cases |
| 5.1 transaction duplicates/conflicts/invalid candidates; physical winner | I | `classification.py`, `quality.py:_verify_policy` | S09, combined conflict/invalid reasons, branch-key and first-physical representative assertions |
| 5.1 product line identity and complete transaction quarantine | I | `product_classification.py`, `quality.py` | S17; missing line/unreadable transaction keys; conflicting/invalid/inconsistent metadata and tampered partial-acceptance tests |
| 5.1 empty/header-only/all-rejected outcomes and independent issues | I | `validation.py`, classification modules, `quality.py` | S10/S12; empty workbook tests in both reporting suites; invalid-price amount-quantification cases |
| 5.2 one outcome per occurrence; known/unquantified money conservation | I | `quality.py:reconcile_transactions` | S15; product independent-price ledger tests; intentional corrupted-classification/integrity tests |
| 5.3 local/upstream separation, unknowns, ordered reasons, statuses | I | `engine.py`, `report_model.py`, `publication.py`, `cli.py` | Bundle/standalone comparison tests; CLI status tests; release `test_exit_json_and_saved_status_agree_without_network` |
| 6.1 Decimal parsing, integer arithmetic, bounds, exact final HALF_UP | I | `money.py`, `product_validation.py` | S04/S05/S06; positive/negative ties, 512-character/38-digit boundaries; release full pipelines with all Decimal traps enabled |
| 6.2 transaction/customer denominators, currency/month/branch partitions | I | `aggregation.py` | S03/S04; branch-scoped customer tests, observed months, independent rollup partitions and negative/zero transactions |
| 6.3 product quantities/prices/weighted denominator/ranking/reconciliation | I | `product_aggregation.py` | S16/S18/S19; cross-branch/currency grouping, zero positive denominator, exact weighted prices, ties and summary-to-product conservation |
| 7.1-7.2 seven files, selected master, hashes, real schemas/counts/amounts | I | `bundle_io.py`, `bundle_schema.py`, `bundle.py` | `test_bundle.py`: missing/altered members, manifest requirements, strict JSON, embedded report, reconciliation partitions, re-signed inconsistent masters |
| 7.3-7.4 exclusions, representative references, provenance, upstream review | I | `bundle_schema.py:parse_exclusions`, `bundle.py:verify_master_agreement` | Exclusion source/duplicate/type/code tests; zero-accepted/conflict bundles; actual sibling integration |
| 8.1 exact sheets/tables/scope/raw diagnostics/literal monetary text | I | `report_model.py`, `workbook.py` | Both reporting suites; read-back values/types/all rows; corruption tests; release ignored-sheet and external-link regressions |
| 8.1 Excel text/XML/row limits, formatting structure | I | `workbook.py` | `test_excel_text_limits_fail_without_truncating`, row-limit fault injection, table/filter/frozen-pane assertions; visual limits below |
| 8.2 isolated charts, minor-unit/scale labels, 15-digit boundary, top ten | I | `report_model.py`, `workbook.py` | Currency-isolation/title/axis/helper tests in both reporting suites; chart-boundary/per-chart and top-ten cases; actual bundle read-back |
| 8.3 deterministic logical output and exact lexical ordering | I | Aggregation modules, `bundle_schema.py`, `report_model.py` | Engine permutation/lexical/ranking tests; release repeated-publication plans, saved tables, sheet/chart counts in both modes |
| 9 required CLI/options/version, JSON and exits 0/1/2 | I | `cli.py`, `__main__.py`, `pyproject.toml`, `version.py` | `test_cli.py`, product/bundle CLI tests, release cross-check of saved status/reasons against exit/JSON; packaging checks below |
| 9 safe paths, aliases, read-only source snapshots, no-clobber publication | I | `readers.py`, `bundle_io.py`, `publication.py` | S21: existing target/hard-link/input alias, simulated links, concurrent target, changed source/bundle, failed writer/read-back/link and owned cleanup; release cleanup-failure regressions |
| 11.1 S01-S24 standalone acceptance cases | I | Default tests under `tests/` | `test_transactions.py`/`test_readers_provenance.py` S01-S15; `test_products.py` S16-S20; report/CLI/release tests S21-S24 |
| 11.2 independent synthetic bundle assertions and standalone unknown scope | I | `tests/bundle_fixtures.py`, `tests/test_bundle.py` | Self-contained schema producers, extensive independent tampering cases; no SheetMerge import/subprocess in default bundle setup |
| 11.3 opt-in actual sibling fixtures and structural failure | I | `tests/conftest.py`, `tests/test_sheetmerge_integration.py` | Separate opt-in run: both masters plus deliberate missing-column failure; original sample hashes/mtimes checked |

Implementation paths in the matrix are under `src/salesreport_auto/`; test files
are under `tests/`. No mandatory feature was identified as wholly unimplemented
after these fixes. The following verification gaps prevent treating all I rows
as universally proven or production-ready.

## Remaining gaps, ambiguities and limitations

1. **Direct test gaps:** case-insensitive extension dispatch is implemented with
   `.suffix.lower()` but has no dedicated uppercase-extension assertion. Output
   rejection above 32,767 UTF-16 units is tested; successful publication at exactly
   32,767 units is not separately asserted. Field/schema ordering has assertions
   for representative issues/reasons, not every possible combined warning/group
   combination. These are narrow unverified mandatory boundary behaviors, not
   evidence that they fail.
2. **Environment-dependent verification:** actual symlink creation was skipped;
   real junctions/dangling links and every Windows ACL/disk-full/interruption case
   were not exercised. Guards and filesystem/publication failures are simulated.
   There is no crash-durability/fsync guarantee, adversarial filesystem guarantee,
   or alternate-filesystem matrix. Hard-link support is required and failure is
   closed. If cleanup is denied after commit, a verified report and an owned
   temporary hard link can coexist; a warning identifies the leftover.
3. **Presentation/scale:** formatting code, filters, panes, cell types and chart
   XML/reference/axis labels are inspected/tested. Actual Excel/LibreOffice visual
   rendering and accessibility have not been manually reviewed. Limits are tested
   with reduced budgets, sparse spans and boundary values; no full 100,000-row
   workbook performance benchmark or exhaustive 200-MiB OOXML attack suite ran.
   The specification explicitly limits this to trusted local business files.
4. **Platform/dependency matrix:** only the stated Windows/Python 3.12 environment
   was exercised. The metadata permits newer Python and multiple dependency
   versions; their compatibility has not been established by this audit. New-user
   virtual-environment creation and installed-wheel execution were checked locally;
   a clean-machine online resolver/build-isolation matrix was not run.
5. **Specification lifecycle wording:** `REPORTING_SPEC.md:3` says application not
   implemented, section 1 says the document is the only deliverable, and section 11
   says tests are to be implemented later. These are historical-stage statements,
   contradicted by the current working application. Proposed editorial change:
   mark the contract implemented for 0.1.0 while preserving its business rules.
   The specification was not changed.
6. **Pandas wording:** section 1 says pandas is used for tabular organization;
   section 10 permits its use. Pandas remains a declared dependency but the actual
   pipeline uses standard-library records/sets/integers rather than importing it.
   No financial behavior depends on pandas. Clarify the introductory wording or
   consider the unused dependency in a later dependency review; do not rewrite
   working arithmetic merely to introduce pandas.
7. **CSV lexical ambiguity:** the contract does not pin an RFC-level quote grammar.
   The implementation uses Python `csv.reader(strict=True)`: malformed closing
   quotes/widths fail, but a quote inside an otherwise unquoted field is retained
   as literal punctuation. Clarify whether stricter RFC 4180 rejection is intended
   before treating that spelling as a contract violation. No speculative parser
   rewrite was made.
8. **Release administration:** no license or author/project URLs are declared.
   Those fields are not required for a valid build, but redistribution/license
   policy should be chosen by the owner before wider distribution. No license
   or ownership details were invented. Source remains uncommitted/untracked.
9. **Deliberate scope limits, not missing MVP work:** no tax/gross/net semantics,
   refund linkage, FX conversion, fractional quantities, numeric XLSX business
   inference, arbitrary schemas, original-source replay, authenticity signature,
   deeper audit JSONL/reconciliation CSV decoding or production certification.
   A verified supplied manifest proves internal consistency only. Exact accepted
   sums and accepted product lines do not prove all sales or lines were supplied.

## Focused changes made

- `readers.py`: distinguish stored empty text (which still defines the record
  span) from nonempty extra-column data, including empty shared strings and XML
  indentation outside text nodes. Explicit
  blank headers remain fatal; invalid shared-string references fail structurally.
- `report_model.py`: add worksheet-selection/ignored-sheet metadata to Data Quality,
  as section 3.1 requires, while retaining Summary disclosures.
- `workbook.py`: retain external-link declarations during output read-back so the
  existing external-link rejection actually works. External targets are not opened.
- `publication.py`, `cli.py`: prevent denied staging-file cleanup from replacing
  the original error or changing a committed report into false failure. Precommit
  cleanup failures become error notes; postcommit failures are explicit warnings
  (`ReportResult.cleanup_warning`) without changing financial status/JSON schema.
- `tests/test_release_audit.py`: 29 additional parameterized regression cases:
  complete pipelines under hostile Decimal contexts; XLSX empty cells and headers;
  malformed/encrypted ZIP members; ignored sheets; failures before/after commit;
  network-blocked status agreement; external-link corruption; logical repeatability.
- `MANIFEST.in`: include the full test fixtures/conftest, specification, audit and
  synthetic CSV examples in the source archive. The initial archive omitted the
  test helpers and specification, making its bundled tests unusable.
- `.gitignore`: ignore documented report folders, existing demo report outputs and
  hidden staging workbooks, in addition to environments/build/cache/temp artifacts.
- `README.md`, `pyproject.toml`: add complete Windows venv setup, executable demo
  commands and expected totals, cleanup behavior/build guidance, and a development
  build frontend. Runtime dependencies and public CLI entry point are unchanged.

No change was made to money formulas, accepted-record semantics, duplicate keys,
product quarantine or SheetMerge files. Specification SHA-256 remains
`65AFA00087310EBA83C9EBBE00CEAD88053C8A00110F6DCBFF0BA72138F9A879`.

## Distribution and installed-package evidence

The corrected `dist/salesreport_auto-0.1.0-py3-none-any.whl` contains 29 regular
files: exactly 24 package Python modules plus 5 `.dist-info` metadata/record files.
`dist/salesreport_auto-0.1.0.tar.gz` contains 52 regular files: package/source
metadata, build configuration, README, specification/audit, synthetic demo CSVs,
all 12 test/helper files and `.gitignore`. Both archives are under 150 KiB;
the largest individual member is the 62,126-byte specification. Exact package
and permitted source-path allowlists were checked, along with safe archive paths,
absence of links/binary workbooks/environments/caches/build outputs, and common
credential/private-key token patterns. No private data or credentials were found
in the inspected contents; a pattern scan is not a proof against every secret.

Metadata checks passed for name/version `salesreport-auto`/`0.1.0`, Python
`>=3.12`, Markdown README, `pandas>=2.2,<4`, `openpyxl>=3.1,<4`, development
extras (`pytest`, `ruff`, `build`) and console entry
`salesreport-auto = salesreport_auto.cli:main`. The wheel tag is `py3-none-any`;
this tag describes packaging and does not establish a tested platform matrix.
The wheel was built through the source-distribution round trip. Setuptools emits
informational "no previously-included" messages for excluded directories that
were never selected; these are not build failures.

The extracted corrected source archive ran the complete default suite successfully
without an adjacent SheetMerge checkout: 521 passed, 4 skipped, 0 failed. The
final archives were also checked against the current source/test bytes. No installed
package import was allowed to substitute checkout code in the wheel smoke probe.

For installed-wheel verification, a separate virtual environment was created
offline and the wheel installed with `--no-index --no-deps`. Existing runtime
dependencies were exposed through a test-only `.pth` path; the old environment's
editable-install hooks were not processed. The probe asserted the imported
package resides in the new environment. This verifies installed package contents
and entry points, but does not claim a fresh online dependency installation.
Module and console invocations both passed `--help`, `--version`, and actual
report generation/read-back for each mode. Every repeated output attempt exited
2 with no success JSON and unchanged existing report hash/mtime. Demo inputs were
also checked unchanged. The installed environment's `pip check` passed.

| Included example | Accepted records | Transactions | Net revenue | Status |
| --- | --- | --- | --- | --- |
| `demo/sales.csv` | 7 | 7 | USD `435.75` | completed / exit 0 |
| `demo/products.csv` | 5 lines | 4 | USD `195.00` | completed / exit 0 |

`git check-ignore` confirmed coverage for existing demo reports, `reports/`,
hidden staging files, `.venv/`, `.test-tmp/`, pytest caches, build/dist and
egg-info. The source specification's SHA-256 is unchanged. Generated audit
workbooks, build logs, archive-check scripts and isolated install/test folders
remain only under ignored workspace directories. Nothing was published.
