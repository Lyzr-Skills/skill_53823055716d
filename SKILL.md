---
name: aggregate-tabular-data
description: >-
  Aggregate tabular data — from an Excel, CSV, TSV, or JSON file, or inline CSV/JSON — by
  grouping on one or more columns and computing metrics such as sum, average, count,
  distinct count, min, max, median, and standard deviation. Returns the result as
  structured JSON. Use this whenever a user needs grouped summaries, subtotals,
  pivot-style rollups, or overall totals of tabular data.
license: MIT
---

# Aggregate Tabular Data

Group rows of tabular data — Excel, CSV, TSV, or JSON — and compute aggregate metrics,
returning structured JSON. The logic lives in `aggregate.py` and can be run as a
command-line program or imported as a Python function (`aggregate_data`).

## When to use this skill

Use it when the user asks for grouped summaries of tabular data (Excel, CSV, TSV, or
JSON), for example:

- "Total sales amount and order count per region."
- "Average price by product category."
- "How many distinct customers per store?"
- "Give me the grand total of the revenue column."

## Requirements

Python 3.9+ with these packages (see `requirements.txt`):

```
pip install -r requirements.txt   # pandas, openpyxl, xlrd
```

## How to run

Command line (recommended for the agent's code sandbox):

```bash
# Group by one column, compute two metrics, sort and take the top rows
python aggregate.py --file-path sales.csv --group-by region \
    --metrics "sum:amount,count:*" --sort-by amount_sum --descending --limit 5

# Multiple group-by columns
python aggregate.py --file-path sales.xlsx --sheet-name 0 \
    --group-by "region,product" --metrics "avg:amount,max:quantity"

# Grand total over all rows (omit --group-by)
python aggregate.py --file-path sales.csv --metrics "sum:amount,nunique:customer_id"

# Inline data instead of a file
python aggregate.py --data "cat,val
A,1
A,3
B,10" --group-by cat --metrics "sum:val,avg:val"

# Pick one table from a multi-table JSON "datasets" envelope (e.g. another agent's output)
python aggregate.py --file-path agent_output.json --dataset medical_claims \
    --group-by "PlanID,IncurredMonth" --metrics "sum:MedPaid,sum:RxPaid"

# Also write the results to a macro-enabled .xlsm, preserving the template's VBA
python aggregate.py --file-path sales.csv --group-by region --metrics "sum:amount,count:*" \
    --output-path report.xlsm --template-path macros_template.xlsm --output-sheet Summary

# Or write a plain .xlsx (no template needed)
python aggregate.py --file-path sales.csv --group-by region --metrics "sum:amount" \
    --output-path report.xlsx
```

Python import:

```python
from aggregate import aggregate_data

json_result = aggregate_data(
    group_by=["region"],
    metrics=["sum:amount", "count:*"],
    file_path="sales.csv",
)
```

## Parameters

| Parameter     | Type            | Required | Description |
|---------------|-----------------|----------|-------------|
| `group_by`    | list / CSV text | No       | Columns to group by (e.g. `region,product`). Empty ⇒ one grand-total record. Leading/trailing whitespace in key values is trimmed so `" North"` and `"North"` group together. |
| `metrics`     | list / CSV text | Yes      | Aggregations as `func:column` items (optional `:alias`). See functions below. Use `count:*` for the row count. |
| `file_path`   | string          | One of file_path / data | Path to a `.csv`, `.tsv`, `.xlsx`, `.xlsm`, `.xls`, or `.json` file. |
| `data`        | string          | One of file_path / data | Inline CSV text or JSON (see JSON input shapes below). |
| `data_format` | string          | No       | `auto` (default), `csv`, `tsv`, `json`, or `excel`. |
| `sheet_name`  | string          | No       | Excel sheet name or 0-based index. Default: first sheet. |
| `dataset`     | string          | No       | For a multi-table JSON *datasets* envelope, which table to aggregate — a dataset name or 0-based index. Empty auto-selects when there is only one dataset. |
| `sort_by`     | string          | No       | Output column to sort by (a group column or a metric alias such as `amount_sum`). |
| `descending`  | boolean         | No       | Sort descending when true. Default false. |
| `limit`       | integer         | No       | Keep only the first N result rows. `0` (default) keeps all. |
| `output_path` | string          | No       | Optional `.xlsx`/`.xlsm` file to also write the result rows to. Adds an `output_file` object to the JSON. |
| `template_path` | string        | No       | Existing macro-enabled `.xlsm` whose VBA is preserved. **Required when `output_path` ends in `.xlsm`.** Optional for `.xlsx` (keeps the template's other sheets). |
| `output_sheet` | string         | No       | Worksheet to (re)write with the results. Default `Aggregation`. Other sheets in a template are left untouched. |
| `include_fingerprint` | boolean | No       | When true, add an `input_fingerprint` block (row count + per-column digests) to the result for troubleshooting drift between runs. Default false. |

**Supported metric functions:** `sum`, `mean`/`avg`, `min`, `max`, `count`,
`nunique`/`distinct`, `median`, `std`, `var`, `first`, `last`.

**Metric alias rule:** each metric is named `{column}_{func}` (e.g. `amount_sum`), or
`count` for `count:*`. Override with a third part, e.g. `sum:amount:total_sales`.

**Numeric parsing:** for numeric metrics (`sum`, `mean`/`avg`, `median`, `std`, `var`) the
target column is coerced to numbers. Values supplied as formatted text are recovered
automatically — currency symbols and thousands separators (`"$1,234.50"`, `"1,000"`),
surrounding whitespace, and accounting-style negatives (`"(250)"` → `-250`). A value that
is still not a number (e.g. `"N/A"`, or an ambiguous decimal comma like `"1,5"`) is
**excluded** from that metric and reported in a `warnings` array (see Output) rather than
silently dropped — so a total is never quietly too low. Blank/`null` cells are treated as
no value and are not reported. If a metric column has **no numeric values at all** (e.g.
you aggregated a money column that is only populated in a *different* dataset), the metric
comes out empty/zero and a `no_numeric_values` warning is emitted — a strong hint that the
wrong `dataset` or column was selected.

## JSON input shapes

JSON is accepted both as a `.json` file (`file_path`) and inline (`data`). The following
shapes are all recognized automatically:

- **Array of records:** `[{"region": "North", "amount": 100}, ...]`
- **Single record object:** `{"region": "North", "amount": 100}`
- **Records wrapped under a key:** `{"data": [ ... ]}` (also `records`, `rows`, `items`,
  `results`, `values`).
- **Column-oriented:** `{"region": ["North", "South"], "amount": [100, 200]}`
- **Nested objects:** flattened into dotted columns, e.g. `{"order": {"region": "North"}}`
  becomes the column `order.region` (group by `order.region`).
- **JSON Lines (NDJSON):** one JSON object per line.
- **"Datasets" envelope (headers + rows):** one or more tables, each given as a `headers`
  list plus `rows` (a list of value lists), optionally nested under wrapper keys — e.g.
  `{"response": {"result": {"datasets": [{"name": "claims", "headers": [...], "rows": [[...], ...]}]}}}`.
  This is the shape emitted by some upstream agents. With more than one dataset, pass
  `dataset` (a name or 0-based index) to choose which to aggregate; the selected name is
  echoed back under `source.dataset`.

## Output

On success (JSON):

```json
{
  "status": "success",
  "source": { "type": "file", "location": "sales.csv", "format": "csv", "sheet": null, "dataset": null },
  "rows_read": 8,
  "columns": ["region", "product", "amount", "quantity", "customer_id"],
  "group_by": ["region"],
  "metrics": ["sum:amount", "count:*"],
  "record_count": 3,
  "records": [
    { "region": "South", "amount_sum": 511.5, "count": 3 },
    { "region": "West",  "amount_sum": 500.0, "count": 2 },
    { "region": "North", "amount_sum": 420.5, "count": 3 }
  ]
}
```

When `output_path` is set, an `output_file` object is added and the same records are
written to the spreadsheet:

```json
{
  "status": "success",
  "record_count": 3,
  "records": [ /* ... */ ],
  "output_file": {
    "path": "C:/reports/report.xlsm",
    "format": "xlsm",
    "sheet": "Summary",
    "row_count": 3,
    "macros_preserved": true,
    "template": "C:/templates/macros_template.xlsm"
  }
}
```

> **Note on `.xlsm`:** macros cannot be generated from scratch — a `template_path`
> pointing at an existing macro-enabled workbook is required, and its VBA is copied into
> the output. For a data-only export with no macros, use an `.xlsx` `output_path`.

When a numeric metric column contains values that could not be parsed as numbers (after
recovering currency/thousands formatting), those cells are excluded and an optional
`warnings` array is added so the exclusion is visible:

```json
{
  "status": "success",
  "records": [ /* ... */ ],
  "warnings": [
    {
      "type": "non_numeric_values_excluded",
      "column": "RxPaid",
      "affected_metrics": ["RxPaid_sum"],
      "excluded_count": 2,
      "sample_values": ["N/A", "pending"],
      "message": "2 value(s) in column 'RxPaid' are not numeric and were excluded from RxPaid_sum. Examples: ['N/A', 'pending']."
    }
  ]
}
```

A `no_numeric_values` warning is emitted instead when the column has **no** numeric values
at all (empty/zero metric), which usually means the wrong `dataset` or column was chosen:

```json
{
  "type": "no_numeric_values",
  "column": "MedPaid",
  "affected_metrics": ["MedPaid_sum"],
  "row_count": 32,
  "message": "Column 'MedPaid' has no numeric values across 32 row(s), so MedPaid_sum is empty/zero. This often means the wrong dataset or column was selected."
}
```

An `inconsistent_group_keys` warning is emitted when a `group_by` column contains values
that differ only by **type or letter case** (e.g. the text `"1"` and the number `1`, or
`"North"` and `"north"`). Grouping is exact, so these split into separate groups and each
group's total can look too low — normalize the column upstream for a stable grouping.
(Leading/trailing **whitespace** is *not* warned about: it is always trimmed from group
keys first, so `" North"`, `"North "` and `"North"` are treated as one group.)

```json
{
  "type": "inconsistent_group_keys",
  "column": "PlanID",
  "distinct_before": 5,
  "distinct_after_normalization": 4,
  "examples": [["'1'", "1"]],
  "message": "Group-by column 'PlanID' has values that differ only by type or letter case and were split into separate groups (e.g. [\"'1'\", '1']). Totals per group may look too low. Normalize the column upstream for a stable grouping."
}
```

The `warnings` key is present only when something was excluded or empty; a clean run omits
it.

### Troubleshooting result drift (`input_fingerprint`)

The aggregation is exact and deterministic: the same input plus the same parameters always
yields the same numbers. So if a total looks different between two runs, either the
parameters differed or **the input data changed**. To tell which, set
`include_fingerprint: true` (CLI: `--include-fingerprint`) and an `input_fingerprint`
object is added to the response:

```json
{
  "status": "success",
  "records": [ /* ... */ ],
  "input_fingerprint": {
    "rows": 40,
    "digest": "3f9a1c7e5b2d0a84",
    "columns": {
      "IncurredMonth": { "digest": "a1b2c3d4e5f60718", "non_null": 40, "null": 0 },
      "PlanID":        { "digest": "9f8e7d6c5b4a3928", "non_null": 40, "null": 0 },
      "MedPaid":       { "digest": "1122334455667788", "non_null": 40, "null": 0 }
    }
  }
}
```

How to read it:

- `rows` — how many rows were aggregated.
- `digest` — an order-insensitive hash over the columns that feed the result (the
  `group_by` keys and the metric source columns).
- `columns` — a per-column digest plus non-null/null counts, to localize *which* column
  changed.

The digests are **order-insensitive** (row order does not matter) and
**value-canonicalized**: they are computed from the *effective* values the aggregation
uses — after whitespace is trimmed from group keys and after numeric coercion — so purely
cosmetic differences (row reordering, `"2,000"` vs `2000`, `"North "` vs `"North"`) do
**not** change a digest. Interpreting a comparison of two runs:

| Same digest? | Same total? | Meaning |
|--------------|-------------|---------|
| yes          | yes         | Reproducible — identical effective data. |
| yes          | no          | Should not happen (the aggregation is deterministic) — report it. |
| no           | no          | The input data changed; the per-column digests show which column. |
| no           | yes         | Data changed but the total happened to be unaffected. |

A differing `digest` means the input the skill received was different — genuine data drift
(e.g. a live upstream source updated between runs), a different slice/window, or a
non-cosmetic value change. It is **not** a defect in the aggregation.

The `input_fingerprint` key is present only when `include_fingerprint` is true; a normal
run omits it.

On failure (JSON, never a stack trace):

```json
{ "status": "error", "error_type": "InvalidColumn", "message": "Column 'foo' not found. Available columns: [...]." }
```

Error types include: `NoInput`, `FileNotFound`, `ParseError`, `EmptyData`, `NoMetrics`,
`UnknownFunction`, `InvalidColumn`, `DuplicateColumn`, `BadMetric`, `InvalidSort`,
`AmbiguousDataset`, `DatasetNotFound`, `UnexpectedError`, and (for Excel output)
`UnsupportedOutput`, `MissingTemplate`, `TemplateNotFound`, `WriteError`,
`MissingDependency`.

## Files

- `aggregate.py` — the skill (CLI + `aggregate_data` function).
- `requirements.txt` — Python dependencies.
- `sample_data.csv` — small dataset for trying the examples above.
