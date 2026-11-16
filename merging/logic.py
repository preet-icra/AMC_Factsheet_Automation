import re
import time
from collections import defaultdict
from pathlib import Path

import pythoncom
import win32com.client


class PPTMergeError(Exception):
    """Raised when the PowerPoint merge cannot be completed."""


class PPTMerger:
    """
    Match the complete text in the PowerPoint object named FundDetails.

    In Equity_FundPages.pptx, two consecutive slides with the same
    FundDetails value are treated as one fund section. The matching slide(s)
    from Equity_Additional_Disclosure.pptx are inserted after the second slide.
    """

    PP_SAVE_AS_OPEN_XML_PRESENTATION = 24
    MSO_GROUP = 6
    FUND_OBJECT_NAME = "object 37"
    REQUIRED_MAIN_SLIDES_PER_FUND = 2

    def __init__(self, progress_callback=None):
        self.progress_callback = progress_callback

    def log(self, message):
        if self.progress_callback:
            self.progress_callback(message)
        else:
            print(message, flush=True)

    @classmethod
    def is_fund_details_name(cls, shape_name):
        """Accept FundDetails and PowerPoint duplicate suffixes."""
        return bool(
            re.fullmatch(
                r"Object 37(?:[ _-]*\d+)?",
                str(shape_name or "").strip(),
                re.IGNORECASE,
            )
        )

    @classmethod
    def find_named_shape_recursive(cls, shape):
        try:
            if cls.is_fund_details_name(shape.Name):
                return shape
        except Exception:
            pass

        try:
            if shape.Type == cls.MSO_GROUP:
                for index in range(1, shape.GroupItems.Count + 1):
                    found = cls.find_named_shape_recursive(
                        shape.GroupItems.Item(index)
                    )
                    if found is not None:
                        return found
        except Exception:
            pass

        return None

    @classmethod
    def find_fund_details_shape(cls, slide):
        for index in range(1, slide.Shapes.Count + 1):
            found = cls.find_named_shape_recursive(slide.Shapes.Item(index))
            if found is not None:
                return found
        return None

    @classmethod
    def extract_all_text(cls, shape):
        """Extract all text inside the FundDetails object."""
        parts = []

        try:
            if shape.Type == cls.MSO_GROUP:
                for index in range(1, shape.GroupItems.Count + 1):
                    value = cls.extract_all_text(shape.GroupItems.Item(index))
                    if value:
                        parts.append(value)
        except Exception:
            pass

        try:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                value = shape.TextFrame.TextRange.Text
                if value:
                    parts.append(str(value))
        except Exception:
            pass

        try:
            if shape.HasTable:
                table = shape.Table
                for row in range(1, table.Rows.Count + 1):
                    for column in range(1, table.Columns.Count + 1):
                        value = cls.extract_all_text(
                            table.Cell(row, column).Shape
                        )
                        if value:
                            parts.append(value)
        except Exception:
            pass

        return "\n".join(parts)

    @staticmethod
    def make_word_key(text):
        """
        Compare every word in order while ignoring case, punctuation and line
        breaks. '&' is normalized to 'AND'.
        """
        if not text:
            return ""

        normalized = str(text).replace("&", " AND ").upper()
        words = re.findall(r"[^\W_]+", normalized, flags=re.UNICODE)
        return "\u241f".join(words)

    @staticmethod
    def display_name_from_key(key):
        return " ".join(key.split("\u241f")) if key else "Unknown fund"

    @classmethod
    def read_slide_record(cls, slide, slide_number):
        shape = cls.find_fund_details_shape(slide)
        if shape is None:
            return None

        text = cls.extract_all_text(shape)
        key = cls.make_word_key(text)
        if not key:
            return None

        return {
            "slide_number": slide_number,
            "shape_name": str(shape.Name),
            "text": text,
            "key": key,
            "fund_name": cls.display_name_from_key(key),
        }

    def scan_presentation(self, presentation, source_label):
        records = []
        total = presentation.Slides.Count

        for slide_number in range(1, total + 1):
            self.log(
                f"Checking {source_label} slide {slide_number} of {total}..."
            )

            record = self.read_slide_record(
                presentation.Slides.Item(slide_number),
                slide_number,
            )

            if record is None:
                self.log("  FundDetails object not found.")
                continue

            records.append(record)
            self.log(f"  Found FundDetails: {record['fund_name']}")

        return records

    @classmethod
    def build_main_sections(cls, main_records):
        """
        Group consecutive slides with the same FundDetails key.

        The expected valid group is exactly two slides, for example:
        slide 1 + slide 2 = AXIS LARGE CAP FUND.
        """
        if not main_records:
            return []

        by_slide = {
            record["slide_number"]: record
            for record in main_records
        }

        sections = []
        used_slides = set()

        for record in sorted(main_records, key=lambda item: item["slide_number"]):
            first_slide = record["slide_number"]

            if first_slide in used_slides:
                continue

            second_slide = first_slide + 1
            second_record = by_slide.get(second_slide)

            if second_record is None:
                raise PPTMergeError(
                    f"{record['fund_name']} starts on slide {first_slide}, but "
                    f"slide {second_slide} has no FundDetails object. Each fund "
                    "must have two consecutive main slides."
                )

            if second_record["key"] != record["key"]:
                raise PPTMergeError(
                    f"FundDetails mismatch between main slides {first_slide} "
                    f"and {second_slide}.\n"
                    f"Slide {first_slide}: {record['fund_name']}\n"
                    f"Slide {second_slide}: {second_record['fund_name']}"
                )

            sections.append(
                {
                    "fund_name": record["fund_name"],
                    "key": record["key"],
                    "first_slide": first_slide,
                    "second_slide": second_slide,
                }
            )

            used_slides.add(first_slide)
            used_slides.add(second_slide)

        return sections

    @staticmethod
    def build_disclosure_map(disclosure_records):
        """Map each exact FundDetails value to its disclosure source slides."""
        mapping = defaultdict(list)
        for record in disclosure_records:
            mapping[record["key"]].append(record["slide_number"])
        return dict(mapping)

    @staticmethod
    def write_report(report_path, sections, disclosure_map):
        lines = [
            "FundDetails Matching Report",
            "=" * 90,
            "Disclosure slides are inserted after slide 2 of each main fund.",
            "",
        ]

        for section in sections:
            sources = disclosure_map.get(section["key"], [])
            lines.append(
                f"{section['fund_name']} | Main slides "
                f"{section['first_slide']}-{section['second_slide']} | "
                f"Disclosure source slides: {sources or 'NO MATCH'}"
            )

        report_path.write_text("\n".join(lines), encoding="utf-8")

    def merge_presentations(self, fund_pages_ppt, disclosure_ppt, output_ppt):
        main_path = Path(fund_pages_ppt).expanduser().resolve()
        disclosure_path = Path(disclosure_ppt).expanduser().resolve()
        output_path = Path(output_ppt).expanduser().resolve()
        report_path = output_path.with_name("FundDetails_Matching_Report.txt")

        if not main_path.is_file():
            raise FileNotFoundError(f"Equity Fund Pages PPT not found:\n{main_path}")

        if not disclosure_path.is_file():
            raise FileNotFoundError(
                f"Additional Disclosure PPT not found:\n{disclosure_path}"
            )

        if main_path == disclosure_path:
            raise PPTMergeError("The two input files must be different.")

        if output_path in {main_path, disclosure_path}:
            raise PPTMergeError("The output file must differ from both inputs.")

        output_path.parent.mkdir(parents=True, exist_ok=True)

        for existing_file in (output_path, report_path):
            if existing_file.exists():
                try:
                    existing_file.unlink()
                except PermissionError as error:
                    raise PPTMergeError(
                        f"Close this file before running again:\n{existing_file}"
                    ) from error

        powerpoint = None
        main_presentation = None
        disclosure_presentation = None
        inserted_total = 0
        matched_funds = []
        unmatched_funds = []

        pythoncom.CoInitialize()

        try:
            self.log("Starting Microsoft PowerPoint...")
            powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
            powerpoint.Visible = True

            try:
                powerpoint.DisplayAlerts = 0
            except Exception:
                pass

            self.log("Opening Equity Fund Pages presentation...")
            main_presentation = powerpoint.Presentations.Open(
                str(main_path), False, False, False
            )

            self.log("Opening Additional Disclosure presentation...")
            disclosure_presentation = powerpoint.Presentations.Open(
                str(disclosure_path), True, False, False
            )

            main_records = self.scan_presentation(
                main_presentation,
                "Equity Fund Pages",
            )

            disclosure_records = self.scan_presentation(
                disclosure_presentation,
                "Additional Disclosure",
            )

            if not main_records:
                raise PPTMergeError(
                    "No FundDetails objects were found in Equity Fund Pages."
                )

            if not disclosure_records:
                raise PPTMergeError(
                    "No FundDetails objects were found in Additional Disclosure."
                )

            main_sections = self.build_main_sections(main_records)
            disclosure_map = self.build_disclosure_map(disclosure_records)

            self.write_report(report_path, main_sections, disclosure_map)
            self.log(f"Matching report created: {report_path}")

            # Process last fund first so original main slide numbers remain valid.
            for section in reversed(main_sections):
                source_slides = disclosure_map.get(section["key"], [])

                if not source_slides:
                    unmatched_funds.append(section["fund_name"])
                    self.log(
                        f"No disclosure match for {section['fund_name']} "
                        f"(main slides {section['first_slide']}-"
                        f"{section['second_slide']})."
                    )
                    continue

                inserted_for_fund = 0

                for source_slide in source_slides:
                    # First disclosure goes directly after the second main slide.
                    # Extra matching disclosure slides follow the first disclosure.
                    insert_after = (
                        section["second_slide"] + inserted_for_fund
                    )

                    self.log(
                        f"Inserting disclosure source slide {source_slide} after "
                        f"slide {insert_after} for {section['fund_name']}..."
                    )

                    before_count = main_presentation.Slides.Count

                    main_presentation.Slides.InsertFromFile(
                        str(disclosure_path),
                        insert_after,
                        source_slide,
                        source_slide,
                    )

                    time.sleep(0.15)
                    inserted = main_presentation.Slides.Count - before_count

                    if inserted != 1:
                        raise PPTMergeError(
                            f"PowerPoint failed to insert disclosure slide "
                            f"{source_slide} for {section['fund_name']}."
                        )

                    inserted_for_fund += 1
                    inserted_total += 1

                matched_funds.append(
                    {
                        "fund_name": section["fund_name"],
                        "main_slides": [
                            section["first_slide"],
                            section["second_slide"],
                        ],
                        "disclosure_source_slides": list(source_slides),
                        "inserted_count": inserted_for_fund,
                    }
                )

                self.log(
                    f"Added {inserted_for_fund} disclosure slide(s) after "
                    f"the second slide of {section['fund_name']}."
                )

            if inserted_total == 0:
                raise PPTMergeError(
                    "No disclosure slides were inserted. Review the report:\n"
                    f"{report_path}"
                )

            self.log("Saving merged presentation...")
            main_presentation.SaveAs(
                str(output_path),
                self.PP_SAVE_AS_OPEN_XML_PRESENTATION,
            )

            time.sleep(1)

            if not output_path.is_file() or output_path.stat().st_size == 0:
                raise PPTMergeError(
                    "PowerPoint did not create a valid merged presentation."
                )

            self.log(f"Merge completed successfully: {output_path}")

            return {
                "output_file": str(output_path),
                "report_file": str(report_path),
                "inserted_slide_count": inserted_total,
                "matched_funds": list(reversed(matched_funds)),
                "unmatched_funds": list(reversed(unmatched_funds)),
            }

        except Exception as error:
            if isinstance(error, (PPTMergeError, FileNotFoundError)):
                raise

            raise PPTMergeError(
                f"PowerPoint merge failed:\n{error}"
            ) from error

        finally:
            if disclosure_presentation is not None:
                try:
                    disclosure_presentation.Close()
                except Exception:
                    pass

            if main_presentation is not None:
                try:
                    main_presentation.Close()
                except Exception:
                    pass

            if powerpoint is not None:
                try:
                    powerpoint.Quit()
                except Exception:
                    pass

            pythoncom.CoUninitialize()


FundPPTMerger = PPTMerger
