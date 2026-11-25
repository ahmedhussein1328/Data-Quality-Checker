# Data Quality Checker

A Python-based desktop application for analyzing and validating CSV and Excel datasets. It helps identify common data quality issues, including duplicate records, missing values, and invalid data types or formats, then generates structured Excel reports.

Built with **Python, Tkinter, Pandas, OpenPyXL, and xlrd**.

## Features

### 1. Modern Desktop Dashboard
- User-friendly desktop interface built with Tkinter and ttk.
- Dark navy sidebar navigation.
- Dedicated sections for Dashboard, Data Preview, Validation Results, Reports, and Application Info.
- Dashboard cards displaying key dataset statistics.

### 2. CSV & Excel Support
- Supports `.csv`, `.xlsx`, and `.xls` files.
- Automatically detects the file extension and selects the appropriate Pandas reader.
- Handles unsupported, empty, and corrupted files with clear error messages.

### 3. Data Preview
- Displays loaded records in a scrollable `ttk.Treeview` table.
- Supports configurable preview limits from 100 to 2,500 rows.
- Makes it easier to inspect datasets before running validation.

### 4. Data Quality Validation

The application performs three main validation processes:

**Duplicate Detection** (`src/duplicate_checker.py`)
- Identifies exact duplicate rows across all columns.

**Missing Values Detection** (`src/missing_checker.py`)
- Detects missing or null values.
- Calculates the number of missing values in each column.
- Identifies rows containing missing values.

**Unified Data Validation** (`src/data_validator.py`)
- Combines data type and format validation into a single module.
- Prevents duplicate issue reporting.
- Handles missing values separately through the Missing Values Checker.
- Validates the following fields:

| Column | Expected Validation |
|---|---|
| `customer_id` | Integer |
| `name` | String |
| `age` | Non-negative integer |
| `email` | String |
| `join_date` | Valid date in `YYYY-MM-DD` format |

### 5. Excel Report Generation
- Generates structured multi-sheet Excel reports.
- Includes six worksheets:
  1. `Summary`
  2. `Original Data`
  3. `Duplicates`
  4. `Missing Counts`
  5. `Missing Rows`
  6. `Data Validation`
- Allows users to export reports to a chosen location.
- Handles locked or open report files gracefully.

### 6. Optional CLI Mode
- Supports command-line execution in addition to the desktop GUI.
- Provides an alternative way to run data quality checks.

---

## Project Structure

```text
Data-Quality-Checker/
│
├── data/
│   ├── customers.csv
│   └── test_customers.csv
│
├── reports/
│   └── data_quality_report.xlsx
│
├── src/
│   ├── __init__.py
│   ├── csv_reader.py
│   ├── data_validator.py
│   ├── duplicate_checker.py
│   ├── missing_checker.py
│   ├── report_generator.py
│   └── gui.py
│
├── main.py
├── requirements.txt
├── README.md
└── .gitignore
```

---

## Technologies Used

| Technology | Purpose |
|---|---|
| Python | Main programming language |
| Pandas | Data reading, processing, and analysis |
| Tkinter / ttk | Desktop graphical user interface |
| OpenPyXL | Excel report generation |
| xlrd | Reading legacy `.xls` Excel files |

---

## Installation

### 1. Clone the Repository

```bash
git clone <repository-url>
cd Data-Quality-Checker
```

### 2. Create a Virtual Environment

```powershell
python -m venv venv
```

### 3. Activate the Virtual Environment

```powershell
.\venv\Scripts\Activate.ps1
```

### 4. Install Dependencies

```bash
pip install -r requirements.txt
```

---

## Running the Application

### Desktop GUI (Default)

```powershell
python main.py
```

### Command-Line Mode (Optional)

```powershell
python main.py --cli data/customers.csv
```

You can also provide an Excel file:

```powershell
python main.py --cli "path\to\dataset.xlsx"
```

---

## Application Workflow

1. Launch the application.
2. Click **Browse File** to select a CSV or Excel dataset.
3. Navigate to **Data Preview** to inspect the loaded records.
4. Click **Run Quality Check** to start the validation process.
5. Review the detected issues in **Validation Results**.
6. Navigate to **Reports** and click **Export Excel Report**.
7. Save the generated Excel report to your preferred location.

---

## Project Goal

The main goal of this project is to simplify and automate basic data quality checks before datasets are used in data analysis or further data engineering workflows.

It provides a practical way to detect common data issues, improve data reliability, and generate organized validation reports.

---

**Developed as a practical Data Engineering project using Python.**