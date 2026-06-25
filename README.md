# Timeline CLI

A standalone command-line tool for generating a timeline image from an Excel workbook or CSV file.

The tool reads timestamped activity rows, sorts them chronologically, and renders a vertical incident-style timeline as PNG or SVG. It can also export the parsed timeline data as JSON.

## Features

- Reads `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, and `.csv` files
- Generates PNG or SVG timeline images
- Supports multi-page output for large timelines
- Exports parsed timeline events to JSON
- Handles ISO-8601 / Zulu timestamps such as `2026-06-23T00:31:53.2971736Z`
- Strips fractional seconds from displayed timestamps
- Optional row filtering using a configurable column
- Useful diagnostics for sheets, columns, skipped rows, and parsing errors

## Requirements

Python 3.9 or later is recommended.

Install dependencies:

```bash
pip install matplotlib openpyxl
```

`openpyxl` is only required for Excel input. CSV input only requires `matplotlib`.

## Expected Input Columns

By default, the tool expects the following columns:

| Column | Purpose |
|---|---|
| `Timestamp_UTC_0` | Event timestamp |
| `Activity` | Event description |
| `MITRE Tactic` | Event category or tactic label |
| `Visualize` | Optional filter column |

By default, only rows where `Visualize` equals `yes` are included.

You can override all column names with command-line options.

## Basic Usage

Generate a PNG timeline:

```bash
python timeline_cli.py sample.xlsx -o timeline.png
```

Generate an SVG timeline:

```bash
python timeline_cli.py sample.xlsx -o timeline.svg
```

Or specify the format explicitly:

```bash
python timeline_cli.py sample.xlsx -o timeline_output --format svg
```

## CSV Input

```bash
python timeline_cli.py timeline.csv -o timeline.png --no-visualize-filter
```

## Custom Sheet and Columns

```bash
python timeline_cli.py sample.xlsx \
  --sheet Timeline \
  --timestamp-col Timestamp_UTC_0 \
  --activity-col Activity \
  --tactic-col "MITRE Tactic" \
  --visualize-col Visualize \
  --visualize-value yes \
  -o timeline.png
```

## Include All Rows

To ignore the filter column and include all rows:

```bash
python timeline_cli.py sample.xlsx -o timeline.png --no-visualize-filter
```

## Multiple Filter Values

You can provide multiple accepted filter values:

```bash
python timeline_cli.py sample.xlsx \
  -o timeline.png \
  --visualize-value yes,true,1
```

## Multi-page Output

For large timelines, split output into multiple files:

```bash
python timeline_cli.py sample.xlsx -o timeline.png --page-size 50
```

This creates files such as:

```text
timeline_001.png
timeline_002.png
timeline_003.png
```

SVG paging works the same way:

```bash
python timeline_cli.py sample.xlsx -o timeline.svg --page-size 50
```

## JSON Export

Export the parsed timeline data as JSON as well as rendering an image:

```bash
python timeline_cli.py sample.xlsx -o timeline.png --json timeline.json
```

Create only JSON output:

```bash
python timeline_cli.py sample.xlsx --json timeline.json --json-only
```

Example JSON output:

```json
[
  {
    "timestamp": "2026-06-23T00:31:53",
    "timestamp_display": "2026-06-23 00:31:53",
    "activity": "Suspicious authentication observed",
    "mitre_tactic": "Credential Access",
    "source_row": 12
  }
]
```

## Listing Sheets and Columns

List workbook sheets:

```bash
python timeline_cli.py sample.xlsx --list-sheets
```

List columns in the selected sheet:

```bash
python timeline_cli.py sample.xlsx --sheet Timeline --list-columns
```

List CSV columns:

```bash
python timeline_cli.py timeline.csv --list-columns
```

## Diagnostics

Show row inclusion and skip statistics:

```bash
python timeline_cli.py sample.xlsx -o timeline.png --stats
```

Show detailed row skip warnings:

```bash
python timeline_cli.py sample.xlsx -o timeline.png --stats -v
```

## Timestamp Handling

The tool accepts common workbook timestamp formats, including:

```text
2026-06-23 00:31:53
2026-06-23 00:31:53.297173
2026-06-23 00:31:53:297173
2026-06-23
2026-06-23T00:31:53.2971736Z
2026-06-23T00:31:53Z
2026-06-23T01:31:53+01:00
```

Timezone-aware timestamps are normalised to UTC for sorting and output. Fractional seconds are stripped from display and JSON output.

For example:

```text
Input:  2026-06-23T00:31:53.2971736Z
Output: 2026-06-23 00:31:53
```

## Common Options

| Option | Description |
|---|---|
| `-o, --output` | Output PNG or SVG path |
| `--format` | Output format: `png` or `svg` |
| `--sheet` | Excel sheet name |
| `--timestamp-col` | Timestamp column name |
| `--activity-col` | Activity column name |
| `--tactic-col` | Category or tactic column name |
| `--visualize-col` | Filter column name |
| `--visualize-value` | Accepted filter value or comma-separated values |
| `--no-visualize-filter` | Include all rows |
| `--page-size` | Number of events per output image |
| `--json` | Write parsed timeline data to JSON |
| `--json-only` | Write JSON only, no image |
| `--list-sheets` | List Excel workbook sheets |
| `--list-columns` | List columns and exit |
| `--max-rows` | Maximum number of rows to scan |
| `--stats` | Print row processing statistics |
| `-v, --verbose` | Print detailed warnings |

## Examples

Generate a standard timeline:

```bash
python timeline_cli.py sample.xlsx -o timeline.png
```

Generate SVG output:

```bash
python timeline_cli.py sample.xlsx -o timeline.svg
```

Generate paged SVG output and JSON:

```bash
python timeline_cli.py sample.xlsx \
  -o timeline.svg \
  --page-size 25 \
  --json timeline.json
```

Generate a timeline from a CSV file without filtering:

```bash
python timeline_cli.py timeline.csv \
  -o timeline.png \
  --no-visualize-filter
```

Debug missing or skipped rows:

```bash
python timeline_cli.py sample.xlsx \
  -o timeline.png \
  --stats \
  -v
```

## Troubleshooting

### No timeline events found

Check that:

- The selected sheet is correct
- The expected columns exist
- Required cells are not empty
- The filter column contains the expected value
- The timestamp values are in a supported format

Try:

```bash
python timeline_cli.py sample.xlsx --list-sheets
python timeline_cli.py sample.xlsx --sheet Timeline --list-columns
python timeline_cli.py sample.xlsx -o timeline.png --stats -v
```

You can also bypass the filter column:

```bash
python timeline_cli.py sample.xlsx -o timeline.png --no-visualize-filter
```

### Missing column errors

Use `--list-columns` to confirm exact column names. Column matching is case-sensitive.

### Excel dependency error

Install `openpyxl`:

```bash
pip install openpyxl
```

### Large timeline output is too tall

Use `--page-size`:

```bash
python timeline_cli.py sample.xlsx -o timeline.png --page-size 50
```

## Attribution

Timeline CLI includes code adapted from Kanvas by WithSecureLabs.

Kanvas is licensed under the GNU General Public License version 3.0.
This project is also licensed under the GNU General Public License version 3.0.

Relevant changes include extracting and simplifying the timeline image generation
logic into a standalone command-line tool, removing Kanvas-specific application
dependencies, adding CSV input, SVG output, JSON export, paging, and standalone
timestamp parsing.

Modified: 2026-06-25
