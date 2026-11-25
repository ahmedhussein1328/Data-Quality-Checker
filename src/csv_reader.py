"""Dataset file reader supporting CSV and Excel (.xlsx, .xls) files."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pandas as pd
from openpyxl.utils.exceptions import InvalidFileException


SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xls"}


def read_file(file_path: str | Path) -> pd.DataFrame:
    """Read a CSV or Excel dataset into a pandas DataFrame based on file extension."""
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    if not path.is_file():
        raise ValueError(f"The selected path is not a file: {path}")

    extension = path.suffix.lower()
    if extension not in SUPPORTED_EXTENSIONS:
        raise ValueError(
            f"Unsupported file format '{extension or '(no extension)'}'. "
            "Please select a .csv, .xlsx, or .xls file."
        )

    if extension == ".csv":
        df = _read_csv_file(path)
    elif extension == ".xlsx":
        df = _read_excel_file(path, engine="openpyxl", extension=extension)
    else:
        df = _read_excel_file(path, engine="xlrd", extension=extension)

    if df.empty or len(df.columns) == 0:
        file_kind = "CSV" if extension == ".csv" else "Excel"
        raise ValueError(f"The {file_kind} file contains no data.")

    return df


def _read_csv_file(path: Path) -> pd.DataFrame:
    """Read a CSV file with clear error messages for empty or malformed files."""
    try:
        return pd.read_csv(path)
    except pd.errors.EmptyDataError as exc:
        raise ValueError("The CSV file is empty.") from exc
    except pd.errors.ParserError as exc:
        raise ValueError("Invalid CSV file format.") from exc
    except UnicodeDecodeError as exc:
        raise ValueError(
            "Unable to decode the CSV file. Please ensure it is a valid text CSV file."
        ) from exc


def _read_excel_file(path: Path, engine: str, extension: str) -> pd.DataFrame:
    """Read an Excel (.xlsx or .xls) file with clear error handling."""
    try:
        return pd.read_excel(path, engine=engine)
    except (PermissionError, OSError):
        raise
    except ImportError as exc:
        raise ValueError(
            f"Cannot read '{extension}' files because the required '{engine}' package is not installed."
        ) from exc
    except (zipfile.BadZipFile, InvalidFileException, ValueError) as exc:
        raise ValueError(f"Invalid or corrupted Excel ({extension}) file.") from exc
    except Exception as exc:
        if exc.__class__.__name__ == "XLRDError":
            raise ValueError(f"Invalid or corrupted Excel ({extension}) file.") from exc
        raise ValueError(f"Could not read the Excel file: {exc}") from exc


def read_csv(file_path: str | Path) -> pd.DataFrame:
    """Backwards-compatible alias for `read_file`."""
    return read_file(file_path)
