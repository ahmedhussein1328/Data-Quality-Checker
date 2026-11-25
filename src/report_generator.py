"""Excel report generator for Data Quality Checker."""

from __future__ import annotations

from pathlib import Path

import pandas as pd


def generate_report(
    df: pd.DataFrame,
    duplicate_rows: pd.DataFrame,
    duplicate_count: int,
    missing_count: pd.Series,
    rows_with_missing: pd.DataFrame,
    total_missing: int,
    validation_issues: pd.DataFrame,
    validation_issue_count: int,
    output_path: str | Path,
) -> Path:
    """Generate a multi-sheet Excel report summarizing data quality findings."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    summary = pd.DataFrame({
        "Metric": [
            "Total Rows",
            "Total Columns",
            "Duplicate Rows",
            "Total Missing Values",
            "Data Validation Issues",
        ],
        "Count": [
            len(df),
            len(df.columns),
            int(duplicate_count),
            int(total_missing),
            int(validation_issue_count),
        ],
    })

    missing_counts_df = missing_count.rename_axis("Column").reset_index(
        name="Missing Count"
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False,
        )
        df.to_excel(
            writer,
            sheet_name="Original Data",
            index=False,
        )
        duplicate_rows.to_excel(
            writer,
            sheet_name="Duplicates",
            index=False,
        )
        missing_counts_df.to_excel(
            writer,
            sheet_name="Missing Counts",
            index=False,
        )
        rows_with_missing.to_excel(
            writer,
            sheet_name="Missing Rows",
            index=False,
        )
        validation_issues.to_excel(
            writer,
            sheet_name="Data Validation",
            index=False,
        )

    return output_path
