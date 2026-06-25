#!/usr/bin/env python3
"""
timeline_cli.py

Create a simple Kanvas-style timeline PNG from an Excel workbook or CSV file.

Minimal runtime dependencies for Excel input:
  pip install openpyxl matplotlib

CSV input only needs matplotlib.

Examples:
  python kanvas_timeline_cli.py sample.xlsx -o timeline.png
  python kanvas_timeline_cli.py sample.xlsx --sheet Timeline --max-rows 200 -o timeline.png
  python kanvas_timeline_cli.py timeline.csv -o timeline.png --no-visualize-filter

Default Kanvas-style columns:
  Timestamp_UTC_0
  Activity
  MITRE Tactic
  Visualize
"""

from __future__ import annotations

import argparse
import csv
import logging
import re
import sys
import textwrap
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, List, Optional, Sequence

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt


DEFAULT_SHEET = "Timeline"
DEFAULT_TIMESTAMP_COL = "Timestamp_UTC_0"
DEFAULT_ACTIVITY_COL = "Activity"
DEFAULT_TACTIC_COL = "MITRE Tactic"
DEFAULT_VISUALIZE_COL = "Visualize"
DEFAULT_VISUALIZE_VALUE = "yes"

DEFAULT_WIDTH = 14.0
DEFAULT_MIN_HEIGHT = 8.0
DEFAULT_HEIGHT_PER_EVENT = 0.32
DEFAULT_DPI = 150

LABEL_MAX_CHARS = 30
DESC_DISPLAY_MAX = 80

logger = logging.getLogger("kanvas_timeline_cli")


@dataclass(frozen=True)
class TimelineEvent:
    timestamp: datetime
    activity: str
    tactic: str


def strip_fractional_seconds_text(value: str) -> str:
    """
    Strip fractional seconds from common workbook timestamp strings.

    Examples:
      2026-06-23T00:31:53.2971736Z -> 2026-06-23T00:31:53Z
      2026-06-23 00:31:53.297173  -> 2026-06-23 00:31:53
      2026-06-23 00:31:53:297173  -> 2026-06-23 00:31:53
    """
    text = str(value).strip()

    text = re.sub(
        r"(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2})\.\d+",
        r"\1",
        text,
    )

    text = re.sub(
        r"(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}):\d+",
        r"\1",
        text,
    )

    return text


def parse_workbook_datetime(value: Any) -> datetime:
    """
    Parse workbook timestamp values.

    Supports:
      - Excel/openpyxl datetime cells
      - 2026-06-23 00:31:53
      - 2026-06-23 00:31:53.297173
      - 2026-06-23 00:31:53:297173
      - 2026-06-23
      - 2026-06-23T00:31:53.2971736Z
      - 2026-06-23T00:31:53Z
      - 2026-06-23T01:31:53+01:00

    Returns naive UTC datetime values with microseconds stripped.
    """
    if isinstance(value, datetime):
        dt = value
    else:
        raw = strip_fractional_seconds_text(str(value).strip())

        if not raw:
            raise ValueError("empty timestamp")

        iso_value = raw.replace(" ", "T")

        if iso_value.endswith(("Z", "z")):
            iso_value = iso_value[:-1] + "+00:00"

        try:
            dt = datetime.fromisoformat(iso_value)
        except ValueError:
            fallback_formats = (
                "%Y-%m-%d %H:%M:%S",
                "%Y-%m-%d",
            )

            for fmt in fallback_formats:
                try:
                    dt = datetime.strptime(raw, fmt)
                    break
                except ValueError:
                    continue
            else:
                raise ValueError(f"unsupported timestamp format: {value!r}")

    if dt.tzinfo is not None:
        dt = dt.astimezone(timezone.utc).replace(tzinfo=None)

    return dt.replace(microsecond=0)


def format_timestamp(value: datetime) -> str:
    return value.strftime("%Y-%m-%d %H:%M:%S")


def clean_text(value: Any, max_chars: Optional[int] = None) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\t", " ").replace("\n", " ").replace("\r", " ")
    text = " ".join(text.split())

    if max_chars is not None and len(text) > max_chars:
        return text[: max_chars - 3] + "..."

    return text


def load_rows_from_xlsx(path: Path, sheet_name: str) -> Iterable[dict[str, Any]]:
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise SystemExit(
            "Excel input requires openpyxl. Install it with: pip install openpyxl"
        ) from exc

    workbook = load_workbook(path, read_only=True, data_only=True)

    if sheet_name not in workbook.sheetnames:
        available = ", ".join(workbook.sheetnames)
        raise SystemExit(f"Sheet {sheet_name!r} not found. Available sheets: {available}")

    sheet = workbook[sheet_name]
    rows = sheet.iter_rows(values_only=True)

    try:
        header_row = next(rows)
    except StopIteration:
        return

    headers = [str(h).strip() if h is not None else f"Column {i + 1}" for i, h in enumerate(header_row)]

    for row in rows:
        yield {
            headers[i]: row[i] if i < len(row) else None
            for i in range(len(headers))
        }


def load_rows_from_csv(path: Path) -> Iterable[dict[str, Any]]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        yield from reader


def load_rows(path: Path, sheet_name: str) -> Iterable[dict[str, Any]]:
    suffix = path.suffix.lower()

    if suffix in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        yield from load_rows_from_xlsx(path, sheet_name)
        return

    if suffix == ".csv":
        yield from load_rows_from_csv(path)
        return

    raise SystemExit(f"Unsupported input type {suffix!r}. Use .xlsx/.xlsm or .csv.")


def build_events(
    rows: Iterable[dict[str, Any]],
    timestamp_col: str,
    activity_col: str,
    tactic_col: str,
    visualize_col: Optional[str],
    visualize_value: str,
    use_visualize_filter: bool,
    max_rows: int,
) -> List[TimelineEvent]:
    events: List[TimelineEvent] = []
    skipped = 0

    visualize_value_normalised = visualize_value.strip().lower()

    for row_num, row in enumerate(rows, start=2):
        if row_num > max_rows + 1:
            break

        try:
            if use_visualize_filter and visualize_col:
                raw_visualize = row.get(visualize_col)
                if str(raw_visualize).strip().lower() != visualize_value_normalised:
                    continue

            timestamp_value = row.get(timestamp_col)
            activity_value = row.get(activity_col)
            tactic_value = row.get(tactic_col)

            if not timestamp_value or not activity_value or not tactic_value:
                skipped += 1
                continue

            timestamp = parse_workbook_datetime(timestamp_value)

            events.append(
                TimelineEvent(
                    timestamp=timestamp,
                    activity=clean_text(activity_value),
                    tactic=clean_text(tactic_value),
                )
            )
        except Exception as exc:
            skipped += 1
            logger.warning("Skipping row %s: %s", row_num, exc)

    events.sort(key=lambda event: event.timestamp)

    if skipped:
        logger.info("Skipped %s row(s)", skipped)

    return events


def draw_timeline(
    events: Sequence[TimelineEvent],
    output_path: Path,
    title: str,
    width: float,
    min_height: float,
    height_per_event: float,
    dpi: int,
) -> None:
    if not events:
        raise SystemExit("No timeline events found. Check sheet name, column names, and visualise filter.")

    fig_height = max(min_height, len(events) * height_per_event)
    fig, ax = plt.subplots(figsize=(width, fig_height))

    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    line_x = 0.12
    ax.axvline(x=line_x, ymin=0, ymax=1, color="black", linewidth=2)

    for index, event in enumerate(events):
        y_pos = 1.0 - (index / max(len(events), 1)) * 0.9

        ax.plot(line_x, y_pos, "o", color="blue", markersize=8)

        ax.text(
            line_x - 0.045,
            y_pos,
            format_timestamp(event.timestamp),
            fontsize=8,
            ha="right",
            va="center",
            color="red",
            weight="bold",
            family="sans-serif",
        )

        tactic = clean_text(event.tactic, LABEL_MAX_CHARS)
        ax.add_patch(
            plt.Rectangle(
                (line_x + 0.012, y_pos - 0.012),
                0.28,
                0.024,
                facecolor="darkorange",
                edgecolor="none",
            )
        )
        ax.text(
            line_x + 0.152,
            y_pos,
            tactic,
            fontsize=8,
            ha="center",
            va="center",
            color="white",
            weight="bold",
            family="sans-serif",
        )

        description = clean_text(event.activity, DESC_DISPLAY_MAX)
        wrapped_description = "\n".join(textwrap.wrap(description, width=90)) if description else ""

        ax.text(
            line_x + 0.31,
            y_pos,
            wrapped_description,
            fontsize=8,
            ha="left",
            va="center",
            family="sans-serif",
        )

    ax.set_xlim(-0.08, 1.0)
    ax.set_ylim(0, 1.02)
    ax.axis("off")
    ax.set_title(title, fontsize=14, weight="bold", pad=20, family="sans-serif")

    plt.tight_layout()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    plt.savefig(output_path, format="png", dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create a Kanvas-style timeline PNG from an Excel workbook or CSV file.",
    )

    parser.add_argument("input", type=Path, help="Input .xlsx/.xlsm workbook or .csv file")
    parser.add_argument("-o", "--output", type=Path, required=True, help="Output PNG path")

    parser.add_argument("--sheet", default=DEFAULT_SHEET, help=f"Workbook sheet name, default: {DEFAULT_SHEET}")
    parser.add_argument("--timestamp-col", default=DEFAULT_TIMESTAMP_COL, help=f"Timestamp column, default: {DEFAULT_TIMESTAMP_COL}")
    parser.add_argument("--activity-col", default=DEFAULT_ACTIVITY_COL, help=f"Activity column, default: {DEFAULT_ACTIVITY_COL}")
    parser.add_argument("--tactic-col", default=DEFAULT_TACTIC_COL, help=f"MITRE tactic column, default: {DEFAULT_TACTIC_COL}")
    parser.add_argument("--visualize-col", default=DEFAULT_VISUALIZE_COL, help=f"Visualise/filter column, default: {DEFAULT_VISUALIZE_COL}")
    parser.add_argument("--visualize-value", default=DEFAULT_VISUALIZE_VALUE, help=f"Required visualise value, default: {DEFAULT_VISUALIZE_VALUE}")
    parser.add_argument("--no-visualize-filter", action="store_true", help="Include all rows instead of filtering on the visualise column")

    parser.add_argument("--max-rows", type=int, default=1000, help="Maximum input rows to scan, default: 1000")
    parser.add_argument("--title", default="Timeline Visualization", help="Chart title")
    parser.add_argument("--width", type=float, default=DEFAULT_WIDTH, help=f"Figure width in inches, default: {DEFAULT_WIDTH}")
    parser.add_argument("--min-height", type=float, default=DEFAULT_MIN_HEIGHT, help=f"Minimum figure height in inches, default: {DEFAULT_MIN_HEIGHT}")
    parser.add_argument("--height-per-event", type=float, default=DEFAULT_HEIGHT_PER_EVENT, help=f"Height per event in inches, default: {DEFAULT_HEIGHT_PER_EVENT}")
    parser.add_argument("--dpi", type=int, default=DEFAULT_DPI, help=f"Output DPI, default: {DEFAULT_DPI}")
    parser.add_argument("-v", "--verbose", action="store_true", help="Show row skip warnings")

    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)

    logging.basicConfig(
        level=logging.INFO if args.verbose else logging.WARNING,
        format="%(levelname)s: %(message)s",
    )

    if not args.input.is_file():
        raise SystemExit(f"Input file not found: {args.input}")

    rows = load_rows(args.input, args.sheet)

    events = build_events(
        rows=rows,
        timestamp_col=args.timestamp_col,
        activity_col=args.activity_col,
        tactic_col=args.tactic_col,
        visualize_col=args.visualize_col,
        visualize_value=args.visualize_value,
        use_visualize_filter=not args.no_visualize_filter,
        max_rows=args.max_rows,
    )

    draw_timeline(
        events=events,
        output_path=args.output,
        title=args.title,
        width=args.width,
        min_height=args.min_height,
        height_per_event=args.height_per_event,
        dpi=args.dpi,
    )

    print(f"Wrote {args.output} with {len(events)} event(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
