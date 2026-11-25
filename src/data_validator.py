"""Unified data type and format validation module."""

from __future__ import annotations

import warnings

import pandas as pd


EXPECTED_COLUMN_RULES = {
    "customer_id": {
        "kind": "integer",
        "allow_negative": True,
        "rule_label": "Integer",
    },
    "name": {
        "kind": "string",
        "rule_label": "String",
    },
    "age": {
        "kind": "integer",
        "allow_negative": False,
        "rule_label": "Non-negative integer",
    },
    "email": {
        "kind": "string",
        "rule_label": "String",
    },
    "join_date": {
        "kind": "date",
        "date_format": "%Y-%m-%d",
        "rule_label": "Date (YYYY-MM-DD)",
    },
}

EXPECTED_COLUMNS = tuple(EXPECTED_COLUMN_RULES.keys())

VALIDATION_COLUMNS = [
    "Row Number",
    "Column",
    "Invalid Value",
    "Validation Rule",
    "Issue",
]


def _validate_integer_series(
    series: pd.Series,
    column: str,
    rule_label: str,
    allow_negative: bool = True,
) -> list[dict]:
    """Validate non-missing values in an integer column without duplicate reporting."""
    records: list[dict] = []
    non_missing = series.notna()
    converted = pd.to_numeric(series, errors="coerce")

    for row_idx in series.index[non_missing]:
        raw_value = series.loc[row_idx]
        num_value = converted.loc[row_idx]

        # Numeric parsing can turn bools into numbers in some cases; guard booleans explicitly.
        if isinstance(raw_value, bool) or pd.isna(num_value) or (num_value % 1 != 0):
            records.append({
                "Row Number": int(row_idx) + 2,
                "Column": column,
                "Invalid Value": raw_value,
                "Validation Rule": rule_label,
                "Issue": "Expected a whole integer value",
            })
        elif not allow_negative and num_value < 0:
            records.append({
                "Row Number": int(row_idx) + 2,
                "Column": column,
                "Invalid Value": raw_value,
                "Validation Rule": rule_label,
                "Issue": "Age cannot be negative",
            })

    return records


def _validate_string_series(
    series: pd.Series,
    column: str,
    rule_label: str,
) -> list[dict]:
    """Validate non-missing values in a string column."""
    records: list[dict] = []
    non_missing = series.notna()

    for row_idx in series.index[non_missing]:
        raw_value = series.loc[row_idx]
        if not isinstance(raw_value, str):
            records.append({
                "Row Number": int(row_idx) + 2,
                "Column": column,
                "Invalid Value": raw_value,
                "Validation Rule": rule_label,
                "Issue": "Expected a text string value",
            })

    return records


def _validate_date_series(
    series: pd.Series,
    column: str,
    rule_label: str,
    date_format: str = "%Y-%m-%d",
) -> list[dict]:
    """Validate non-missing date values and YYYY-MM-DD formatting in one pass."""
    records: list[dict] = []
    non_missing = series.notna()

    strict_dates = pd.to_datetime(series, format=date_format, errors="coerce")
    invalid_mask = non_missing & strict_dates.isna()

    if not invalid_mask.any():
        return records

    # Distinguish completely invalid dates from valid dates using another format.
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        loose_dates = pd.to_datetime(series[invalid_mask], errors="coerce")

    for row_idx in series.index[invalid_mask]:
        raw_value = series.loc[row_idx]
        if pd.isna(loose_dates.loc[row_idx]):
            issue_text = f"Invalid date or format (expected {date_format})"
        else:
            issue_text = "Date must use YYYY-MM-DD format"

        records.append({
            "Row Number": int(row_idx) + 2,
            "Column": column,
            "Invalid Value": raw_value,
            "Validation Rule": rule_label,
            "Issue": issue_text,
        })

    return records


def validate_data(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Run unified type and format validation on all matching columns in `df`.

    Missing values are intentionally skipped because they are reported by
    `missing_checker.find_missing_values`. Each invalid cell is reported at most once.

    Returns:
        A tuple of `(issues_df, issue_count)`.
    """
    records: list[dict] = []

    for column, spec in EXPECTED_COLUMN_RULES.items():
        if column not in df.columns:
            continue

        series = df[column]
        kind = spec["kind"]
        rule_label = spec["rule_label"]

        if kind == "integer":
            records.extend(
                _validate_integer_series(
                    series,
                    column=column,
                    rule_label=rule_label,
                    allow_negative=spec.get("allow_negative", True),
                )
            )
        elif kind == "string":
            records.extend(
                _validate_string_series(
                    series,
                    column=column,
                    rule_label=rule_label,
                )
            )
        elif kind == "date":
            records.extend(
                _validate_date_series(
                    series,
                    column=column,
                    rule_label=rule_label,
                    date_format=spec.get("date_format", "%Y-%m-%d"),
                )
            )

    if records:
        issues_df = (
            pd.DataFrame(records, columns=VALIDATION_COLUMNS)
            .sort_values(by=["Row Number", "Column"], kind="stable")
            .reset_index(drop=True)
        )
    else:
        issues_df = pd.DataFrame(columns=VALIDATION_COLUMNS)

    return issues_df, len(issues_df)
