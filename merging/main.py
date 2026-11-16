import sys
import traceback
from pathlib import Path

from logic import PPTMerger, PPTMergeError


def show_progress(message):
    print(message, flush=True)


def main():
    project_folder = Path(__file__).resolve().parent

    fund_pages_file = project_folder / "EquityFundPages.pptx"
    disclosure_file = project_folder / "EquityOutput.pptx"
    output_file = project_folder / "Merged_Equity_Factsheet.pptx"

    print("=" * 78)
    print("FUND DISCLOSURE MERGER")
    print("=" * 78)
    print(f"Fund Pages: {fund_pages_file}")
    print(f"Disclosure: {disclosure_file}")
    print(f"Output: {output_file}")
    print()

    if not fund_pages_file.is_file():
        print(f"ERROR: File not found:\n{fund_pages_file}")
        return 1

    if not disclosure_file.is_file():
        print(f"ERROR: File not found:\n{disclosure_file}")
        return 1

    try:
        merger = PPTMerger(progress_callback=show_progress)

        result = merger.merge_presentations(
            fund_pages_ppt=str(fund_pages_file),
            disclosure_ppt=str(disclosure_file),
            output_ppt=str(output_file),
        )

        print()
        print("=" * 78)
        print("MERGE COMPLETED")
        print("=" * 78)
        print(f"Inserted disclosure slides: {result['inserted_slide_count']}")
        print(f"Merged PPT: {result['output_file']}")
        print(f"Matching report: {result['report_file']}")

        if result["unmatched_funds"]:
            print("\nFunds without matching disclosure slides:")
            for fund_name in result["unmatched_funds"]:
                print(f"- {fund_name}")

        return 0

    except (FileNotFoundError, PPTMergeError) as error:
        print("\nMERGE FAILED")
        print(error)
        return 1

    except Exception as error:
        print("\nUNEXPECTED ERROR")
        print(error)
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
