"""Application entry point for Data Quality Checker."""

import argparse
from pathlib import Path

from src.csv_reader import read_file
from src.data_validator import validate_data
from src.duplicate_checker import find_duplicates
from src.gui import launch_app
from src.missing_checker import find_missing_values
from src.report_generator import generate_report


def run_cli(data_path: Path) -> int:
    """Run the command-line quality-check workflow on a CSV or Excel file."""
    try:
        df = read_file(data_path)
    except (FileNotFoundError, ValueError, PermissionError, OSError) as error:
        print(f"\nError: {error}")
        return 1

    print("\n===== DATA QUALITY CHECKER =====")
    print("\n===== DATASET INFORMATION =====")
    print(f"File: {data_path}")
    print(f"Total Rows: {len(df)}")
    print(f"Total Columns: {len(df.columns)}")
    print("\nDataset Preview:")
    print(df)

    duplicate_rows, duplicate_count = find_duplicates(df)
    print("\n===== DUPLICATE DETECTION =====")
    print(f"Duplicate Rows Found: {duplicate_count}")
    if duplicate_count > 0:
        print("\nDuplicate Records:")
        print(duplicate_rows)
    else:
        print("\nNo Duplicate Records Found.")

    missing_count, rows_with_missing, total_missing = find_missing_values(df)
    print("\n===== MISSING VALUES DETECTION =====")
    print("\nMissing Values Per Column:")
    print(missing_count)
    print(f"\nTotal Missing Values: {total_missing}")
    if total_missing > 0:
        print("\nRows Containing Missing Values:")
        print(rows_with_missing)
    else:
        print("\nNo Missing Values Found.")

    validation_issues, validation_issue_count = validate_data(df)
    print("\n===== DATA VALIDATION (TYPES & FORMATS) =====")
    print(f"Validation Issues Found: {validation_issue_count}")
    if validation_issue_count > 0:
        print(validation_issues.to_string(index=False))
    else:
        print("No Data Validation Issues Found.")

    print("\n===== REPORT GENERATION =====")
    report_path = Path(__file__).parent / "reports" / "data_quality_report.xlsx"
    generated_file = generate_report(
        df,
        duplicate_rows,
        duplicate_count,
        missing_count,
        rows_with_missing,
        total_missing,
        validation_issues,
        validation_issue_count,
        report_path,
    )
    print("\nExcel Report Generated Successfully!")
    print(f"Report Location: {generated_file}")
    return 0


def main() -> None:
    parser = argparse.ArgumentParser(description="Data Quality Checker Desktop Application")
    parser.add_argument(
        "--cli",
        action="store_true",
        help="Run in command-line mode instead of launching the desktop GUI.",
    )
    parser.add_argument(
        "data_file",
        nargs="?",
        default=str(Path(__file__).parent / "data" / "test_customers.csv"),
        help="Optional path to a CSV or Excel (.xlsx, .xls) file when running with --cli.",
    )
    args = parser.parse_args()

    if args.cli:
        raise SystemExit(run_cli(Path(args.data_file)))

    launch_app()


if __name__ == "__main__":
    main()
