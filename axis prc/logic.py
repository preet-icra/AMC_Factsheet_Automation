"""Update PRC values in Debt_Additional_Disclosure.pptx.

The script reads each fund's PRC value from
ALL_Product Labelling_Jun_2026.pptx and writes only that value into the
existing PRC table in Debt_Additional_Disclosure.pptx. It does not replace,
resize, or rebuild the target PRC tables.

Install:
    pip install python-pptx

Run:
    python logic.py

Optional:
    python logic.py --source "ALL_Product Labelling_Jun_2026.pptx" \
                    --target "Debt_Additional_Disclosure.pptx" \
                    --output "Debt_Additional_Disclosure_Updated.pptx"
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from pptx import Presentation
from pptx.enum.text import MSO_VERTICAL_ANCHOR, PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.table import Table
from pptx.util import Cm, Pt

DEFAULT_SOURCE = "ALL_Product Labelling_Aug_2026.pptx"
DEFAULT_TARGET = "Debt_AddtionalD.pptx"
DEFAULT_OUTPUT = "Debt_Additional_Disclosure_Updated.pptx"
DEFAULT_REPORT = "PRC_Update_Summary.txt"

# PRC tables for these funds must be deleted from the target deck.
DELETE_PRC_FUNDS = (
    "Axis Nifty AAA Bond Plus SDL Apr 2026 50:50 ETF",
    "Axis Nifty AAA Bond Plus SDL Apr 2026 50:50 ETF FOF",
    "Axis US Specific Treasury Dynamic Debt Passive FOF",
)

# Accept both Roman and Arabic source forms, for example A-I or A-1.
PRC_RE = re.compile(r"(?<![A-Z0-9])([ABC])\s*[-–—]\s*(III|II|I|3|2|1)(?![A-Z0-9])", re.I)
FUND_RE = re.compile(r"\bAxis\b[^\r\n]*?\bFund\b(?:\s*of\s*Fund)?", re.I)


def compact(text: str) -> str:
    """Standardize whitespace and dash variants without changing wording."""
    text = (text or "").replace("\u00a0", " ")
    text = text.replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", text).strip()


def match_key(text: str) -> str:
    """Case-insensitive, whitespace-normalized exact-match key."""
    return compact(text).casefold()


def shape_text(shape) -> str:
    if getattr(shape, "has_text_frame", False):
        return compact(shape.text)
    if getattr(shape, "has_table", False):
        return compact(" ".join(cell.text for row in shape.table.rows for cell in row.cells))
    return ""


def table_text(table: Table) -> str:
    return compact(" ".join(cell.text for row in table.rows for cell in row.cells))


def is_prc_table(table: Table) -> bool:
    text = table_text(table).casefold()
    return (
        "potential risk class" in text
        and "credit risk" in text
        and "interest rate risk" in text
        and "class a" in text
        and "class b" in text
        and "class c" in text
    )


def extract_prc_value(table: Table) -> Optional[str]:
    """Return the source PRC text exactly, with surrounding whitespace removed."""
    candidates: List[str] = []
    for row in table.rows:
        for cell in row.cells:
            raw = compact(cell.text)
            match = PRC_RE.search(raw.upper())
            if match:
                # Preserve whether the source uses Roman or Arabic notation.
                candidates.append(f"{match.group(1).upper()}-{match.group(2).upper()}")

    unique = list(dict.fromkeys(candidates))
    if len(unique) == 1:
        return unique[0]
    if len(unique) > 1:
        raise ValueError(f"Multiple PRC values found in one table: {unique}")
    return None


def extract_fund_name_from_text(text: str) -> Optional[str]:
    """Extract one Axis fund name from a title or fund-label table."""
    text = compact(text)
    match = FUND_RE.search(text)
    if not match:
        return None

    name = compact(match.group(0))

    # Remove labels that can occur before the actual name in a table cell.
    name = re.sub(r"^(Fund Name\s*&\s*Benchmark\s*)", "", name, flags=re.I)
    return name.strip(" :-")


def fund_name_from_product_table(table: Table) -> Optional[str]:
    """Identify the fund-label table that immediately precedes a PRC table."""
    text = table_text(table)
    low = text.casefold()
    if "fund name" not in low or "benchmark" not in low:
        return None
    return extract_fund_name_from_text(text)


def build_source_prc_map(source: Presentation) -> Tuple[Dict[str, Tuple[str, str]], List[str]]:
    """Pair each Product Labelling fund table with the next PRC table on the slide."""
    mapping: Dict[str, Tuple[str, str]] = {}
    warnings: List[str] = []

    for slide_no, slide in enumerate(source.slides, start=1):
        pending_fund: Optional[str] = None

        # Shape order normally follows the XML/z-order used by the source deck.
        for shape in slide.shapes:
            if not getattr(shape, "has_table", False):
                continue

            table = shape.table
            fund = fund_name_from_product_table(table)
            if fund:
                pending_fund = fund
                continue

            if not is_prc_table(table):
                continue

            value = extract_prc_value(table)
            if not pending_fund:
                warnings.append(f"Source slide {slide_no}: PRC table has no preceding fund table.")
                continue
            if not value:
                warnings.append(f"Source slide {slide_no}: no PRC value found for '{pending_fund}'.")
                pending_fund = None
                continue

            key = match_key(pending_fund)
            if key in mapping and mapping[key][1] != value:
                warnings.append(
                    f"Source slide {slide_no}: duplicate '{pending_fund}' has conflicting "
                    f"values {mapping[key][1]} and {value}; latest value used."
                )
            mapping[key] = (pending_fund, value)
            pending_fund = None

    return mapping, warnings


def target_fund_name(slide) -> Optional[str]:
    """Find the main fund title on a Debt Additional Disclosure slide."""
    candidates: List[Tuple[int, int, str]] = []

    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        text = compact(shape.text)
        fund = extract_fund_name_from_text(text)
        if not fund:
            continue

        # Prefer smaller title-like text boxes near the top of the slide.
        candidates.append((len(text), int(getattr(shape, "top", 0)), fund))

    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[0][2]


def find_target_prc_table(slide) -> Optional[Table]:
    for shape in slide.shapes:
        if getattr(shape, "has_table", False) and is_prc_table(shape.table):
            return shape.table
    return None


def all_slide_text(slide) -> str:
    """Return normalized text from all text-bearing shapes and tables."""
    return compact(" ".join(shape_text(shape) for shape in slide.shapes if shape_text(shape)))


def deletion_fund_for_slide(slide) -> Optional[str]:
    """Return an exact configured deletion fund found on the slide."""
    text_key = match_key(all_slide_text(slide))

    # Test longer names first so the ETF FOF is not mistaken for the ETF.
    for fund in sorted(DELETE_PRC_FUNDS, key=len, reverse=True):
        fund_key = match_key(fund)
        # Boundary check avoids matching a configured name inside a longer name.
        pattern = r"(?<![a-z0-9])" + re.escape(fund_key) + r"(?![a-z0-9])"
        if re.search(pattern, text_key):
            return fund
    return None


def delete_prc_tables(slide) -> int:
    """Delete every PRC table shape on a slide and return the deletion count."""
    shapes_to_delete = [
        shape for shape in slide.shapes
        if getattr(shape, "has_table", False) and is_prc_table(shape.table)
    ]
    for shape in shapes_to_delete:
        shape.element.getparent().remove(shape.element)
    return len(shapes_to_delete)


def prc_position(value: str) -> Tuple[int, int]:
    """Return matrix row/column for A-I, C-II, B-III, A-1, C-2, etc."""
    match = PRC_RE.fullmatch(compact(value).upper())
    if not match:
        raise ValueError(f"Invalid PRC value: {value}")

    credit = match.group(1).upper()
    interest = match.group(2).upper()

    roman_to_number = {"I": 1, "II": 2, "III": 3, "1": 1, "2": 2, "3": 3}
    credit_col = {"A": 1, "B": 2, "C": 3}[credit]

    # PRC table rows: 0 heading, 1/2 column headings, 3 Class I, 4 Class II, 5 Class III.
    risk_row = {1: 3, 2: 4, 3: 5}[roman_to_number[interest]]
    return risk_row, credit_col


def clear_existing_prc_values(table: Table) -> None:
    """Clear old PRC codes and remove highlighting from all empty matrix cells."""
    # PRC matrix data area:
    # rows 3-5 = Class I, II, III
    # columns 1-3 = Class A, B, C
    max_row = min(6, len(table.rows))
    max_col = min(4, len(table.columns))

    for row_idx in range(3, max_row):
        for col_idx in range(1, max_col):
            cell = table.cell(row_idx, col_idx)

            # Remove any previously inserted PRC value.
            if PRC_RE.fullmatch(compact(cell.text).upper()):
                cell.text = ""

            # Remove old highlight from every matrix cell. The selected value
            # cell is highlighted again in format_value_cell().
            cell.fill.background()


def format_value_cell(cell, value: str) -> None:
    """Format only the inserted PRC value, without changing the table itself."""
    cell.text = value

    # Highlight only the PRC-value cell with #F1F1F1.
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor(0xF1, 0xF1, 0xF1)

    cell.margin_top = Cm(0.10)
    cell.margin_bottom = Cm(0.10)
    cell.margin_left = Cm(0.10)
    cell.margin_right = Cm(0.10)
    cell.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

    for paragraph in cell.text_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.name = "Arial"
            run.font.size = Pt(8)
            run.font.bold = False


def write_prc_value(table: Table, source_value: str) -> None:
    clear_existing_prc_values(table)
    row_idx, col_idx = prc_position(source_value)

    if row_idx >= len(table.rows) or col_idx >= len(table.columns):
        raise IndexError(
            f"Target PRC table is {len(table.rows)}x{len(table.columns)}; "
            f"cannot place {source_value} at ({row_idx}, {col_idx})."
        )

    format_value_cell(table.cell(row_idx, col_idx), source_value)


def update_deck(source_path: Path, target_path: Path, output_path: Path, report_path: Path) -> None:
    source_prs = Presentation(str(source_path))
    target_prs = Presentation(str(target_path))
    source_map, warnings = build_source_prc_map(source_prs)

    checked = 0
    matches = 0
    updated = 0
    unmatched: List[str] = []
    no_target_table: List[str] = []
    update_errors: List[str] = []
    deleted_prc_tables = 0
    deleted_funds: List[str] = []
    seen_funds = set()

    for slide_no, slide in enumerate(target_prs.slides, start=1):
        deletion_fund = deletion_fund_for_slide(slide)
        if deletion_fund:
            deleted_count = delete_prc_tables(slide)
            if deleted_count:
                deleted_prc_tables += deleted_count
                if deletion_fund not in deleted_funds:
                    deleted_funds.append(deletion_fund)
                print(f"Deleted {deleted_count} PRC table(s) on slide {slide_no}: {deletion_fund}")
            # Never insert or update a PRC value for these excluded funds.
            continue

        fund = target_fund_name(slide)
        if not fund:
            continue

        key = match_key(fund)

        # A fund can span multiple disclosure slides. Count/check it once, but update
        # the slide that actually contains the PRC table.
        if key not in seen_funds:
            checked += 1
            seen_funds.add(key)
            if key not in source_map:
                unmatched.append(fund)

        if key not in source_map:
            continue

        table = find_target_prc_table(slide)
        if table is None:
            continue

        matches += 1
        source_name, value = source_map[key]
        try:
            write_prc_value(table, value)
            updated += 1
            print(f"Updated slide {slide_no}: {fund} -> {value}")
        except Exception as exc:
            update_errors.append(f"Slide {slide_no}, {fund}: {exc}")

    # Matched source funds in the target that never had a target PRC table.
    for key in sorted(seen_funds):
        if key in source_map:
            source_name, _ = source_map[key]
            found = False
            for slide in target_prs.slides:
                fund = target_fund_name(slide)
                if fund and match_key(fund) == key and find_target_prc_table(slide) is not None:
                    found = True
                    break
            if not found:
                no_target_table.append(source_name)

    target_prs.save(str(output_path))

    lines = [
        "PRC UPDATE SUMMARY",
        "==================",
        f"Total unique fund names checked: {checked}",
        f"Total exact fund matches with target PRC tables: {matches}",
        f"Total PRC values inserted/updated: {updated}",
        f"Total PRC tables deleted: {deleted_prc_tables}",
        "",
        "Funds whose PRC tables were deleted:",
    ]
    lines.extend([f"- {name}" for name in deleted_funds] or ["- None"] )
    lines.extend([
        "",
        "Unmatched fund names:",
    ])
    lines.extend([f"- {name}" for name in unmatched] or ["- None"])
    lines.extend(["", "Matched funds with no PRC table in target:"])
    lines.extend([f"- {name}" for name in no_target_table] or ["- None"])
    lines.extend(["", "Warnings:"])
    lines.extend([f"- {msg}" for msg in warnings] or ["- None"])
    lines.extend(["", "Update errors:"])
    lines.extend([f"- {msg}" for msg in update_errors] or ["- None"])

    report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("\n" + "\n".join(lines))
    print(f"\nSaved PowerPoint: {output_path}")
    print(f"Saved summary: {report_path}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Copy PRC values between the two PowerPoint decks.")
    parser.add_argument("--source", default=DEFAULT_SOURCE)
    parser.add_argument("--target", default=DEFAULT_TARGET)
    parser.add_argument("--output", default=DEFAULT_OUTPUT)
    parser.add_argument("--report", default=DEFAULT_REPORT)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    source = Path(args.source)
    target = Path(args.target)

    if not source.exists():
        raise FileNotFoundError(f"Source file not found: {source.resolve()}")
    if not target.exists():
        raise FileNotFoundError(f"Target file not found: {target.resolve()}")

    update_deck(source, target, Path(args.output), Path(args.report))


if __name__ == "__main__":
    main()
