# Oracle EBS Python Automation

A growing portfolio of Python tools for Oracle E-Business Suite procurement reporting, data-quality monitoring, and repetitive operational tasks.

The repository is organised as a collection of small, independent automation modules. New Python scripts can be added without changing the overall project structure.

## Why This Project

ERP support teams often spend significant time manually:

- extracting and summarising purchasing data;
- identifying incomplete or abnormal transaction records;
- preparing recurring Excel reports and charts;
- repeating configuration or data-entry steps in Oracle EBS.

These tasks can be slow, inconsistent, and difficult to audit. This project explores how Python can turn them into repeatable support workflows.

## Current Modules

| Module | Category | Business Purpose | Main Output | Status |
| --- | --- | --- | --- | --- |
| [`po_headers_report.py`](./po_headers_report.py) | Reporting | Demonstrates purchase-order processing and Excel export using anonymised sample data | Console summary and Excel workbook | Runnable sample |
| [`PO_Count.py`](./PO_Count.py) | Reporting | Analyses Oracle EBS purchase-order volume and approval status for the previous 30 days | Excel workbook with trend and status charts | Oracle environment required |
| [`auto_sql.py`](./auto_sql.py) | Data quality / Application Support | Detects incomplete Oracle EBS purchasing distributions and provides structured execution evidence for troubleshooting | Date-stamped Excel exception report and application log | Oracle environment required |
| [`creat_office.py`](./creat_office.py) | UI automation | Creates account-rule records in Oracle EBS from spreadsheet input | Automated browser data entry | Environment-specific proof of concept |

> When a new script is added, add one row to this table and create a matching section under **Module Details**.

## Module Details

### 1. Sample Procurement Report

**File:** `po_headers_report.py`

Uses sample purchase-order data to demonstrate:

- purchase-order detail processing;
- counts by approval status;
- totals separated by currency;
- Excel detail and summary export.

Run it without an Oracle connection:

```bash
python po_headers_report.py
```

Output:

```text
output/erp_procurement_report.xlsx
```

### 2. Purchase Order Trend Report

**File:** `PO_Count.py`

Queries `PO_HEADERS_ALL` and produces:

- daily purchase-order counts for the previous 30 days;
- purchase-order counts grouped by approval status;
- a line chart for daily volume;
- a bar chart for status distribution.

### 3. Oracle ERP Data Quality Monitor

**File:** `auto_sql.py`

#### Business Problem

Incomplete purchasing-distribution records can create downstream support issues when expected requisition or delivery references are missing.

Manually locating these records is repetitive, and Application Support also needs enough execution evidence to distinguish database, query, data-processing, and report-export failures.

#### Monitoring Use Case

The script queries `PO_DISTRIBUTIONS_ALL` for records where all of the following fields are missing:

- `REQ_DISTRIBUTION_ID`
- `DELIVER_TO_LOCATION_ID`
- `DELIVER_TO_PERSON_ID`

When issues are found, the records are converted into a pandas DataFrame and exported to a date-stamped Excel report for investigation.

#### Application Workflow

```text
Job Start
  ↓
Oracle Client Initialisation
  ↓
Database Connection
  ↓
PO Distribution Data-Quality Query
  ↓
DataFrame Creation
  ↓
Excel Exception Report
  ↓
Connection Cleanup

```

The workflow is decomposed into focused functions so database access, SQL execution, DataFrame construction, report export, and cleanup have separate responsibilities.

#### Exception Handling

The monitor uses application-specific exception boundaries:

- `DatabaseConnectionError` — Oracle Client or database connection failures
- `QueryExecutionError` — Oracle SQL execution failures
- `DataFrameBuildError` — DataFrame construction failures
- `ReportExportError` — Excel or file-system failures

Lower-level functions raise meaningful exceptions while the workflow decides how failures are logged.

Oracle exception chaining is preserved with `raise ... from error`, allowing the application-level failure to remain connected to the underlying technical cause.

The Oracle cursor uses a Context Manager so it is released even when query execution fails.

#### Configuration and Secret Management

Environment-specific Oracle settings are kept outside the application workflow:

```text
.env
  ↓
config.py
  ↓
auto_sql.py
```

`config.py` loads and validates required environment variables while centralising application settings such as output paths and logging configuration.

Required local values are documented in `.env.example`:

```dotenv
ORACLE_USER=YOUR_ORACLE_USERNAME
ORACLE_PASSWORD=YOUR_ORACLE_PASSWORD
ORACLE_DSN=YOUR_HOST:PORT/SERVICE_NAME
ORACLE_CLIENT_PATH=C:\path\to\oracle\instantclient
```

The real `.env` file is excluded from Git. Real credentials, internal connection details, and production ERP data are not intended for this public repository.

#### Logging and Application Support Troubleshooting

The monitor writes structured execution logs to both the console and `logs/application.log`.

Logs include key operational stages such as:

- Oracle Client initialisation
- database connection status
- detected issue count
- report export
- exception details
- database connection cleanup

`TimedRotatingFileHandler` rotates the application log daily and retains 14 backup files to prevent unlimited log growth.

This supports an Application Support troubleshooting workflow based on execution evidence. A support analyst can use the logs to identify:

- the last confirmed successful stage;
- the first failed stage;
- whether Oracle Client and database connectivity can be ruled out;
- whether DataFrame or Excel stages were reached;
- whether connection cleanup completed after a failure.

This module applies production-oriented engineering practices for maintainability and supportability, but it is a portfolio implementation and is **not presented as a deployed production application**.

### 4. Oracle EBS Account-Rule Automation

**File:** `creat_office.py`

Reads login and setup data from Excel, then uses Selenium and PyAutoGUI to automate repeated Oracle EBS account-rule entry.

This module is a proof of concept and currently depends on environment-specific URLs, browser configuration, file paths, and page elements.

## Technology Stack

| Area | Technologies |
| --- | --- |
| Programming | Python, SQL |
| ERP and database | Oracle E-Business Suite, Oracle Database |
| Data processing | pandas |
| Excel reporting | openpyxl |
| Configuration | python-dotenv, environment variables |
| Operational support | Python logging, timed log rotation |
| Database connectivity | cx_Oracle |
| UI automation | Selenium, PyAutoGUI, Pyperclip |

## Repository Structure

```text
python-automation/
├── README.md
├── auto_sql.py
├── config.py
├── requirements.txt
├── .env.example
├── .gitignore
├── po_headers_report.py
├── PO_Count.py
├── creat_office.py
├── output/                 # Generated reports
└── logs/                   # Runtime logs; log files excluded from Git
```

As the project grows, modules can be reorganised into the following structure:

```text
python-automation/
├── README.md
├── requirements.txt
├── .env.example
├── reporting/
├── data_quality/
├── ui_automation/
├── shared/
├── tests/
└── output/
```

The folder migration is planned only when the number of scripts makes the current flat structure difficult to maintain.

## Configuration

`auto_sql.py` separates environment-specific Oracle configuration from application logic.

Create a local `.env` file based on `.env.example`:

```dotenv
ORACLE_USER=your_username
ORACLE_PASSWORD=your_password
ORACLE_DSN=host:port/service_name
ORACLE_CLIENT_PATH=C:\path\to\oracle\instantclient
```

`config.py` loads and validates these values and centralises non-secret application settings including:

- report output directory and filename format;
- Excel engine;
- logging directory and filename;
- logging level and format;
- log rotation interval and backup retention.

The real `.env` file is excluded by `.gitignore`.

Other modules may still require environment-specific configuration such as Oracle EBS URLs, browser settings, or local input paths.

Never commit real credentials, internal URLs, session parameters, or production ERP data.

## Adding a New Automation Module

Use the following checklist whenever a new `.py` file is added:

1. Give the file a clear, task-based name.
2. Keep credentials and machine-specific paths outside the code.
3. Add a module docstring describing the business problem.
4. Add an `if __name__ == "__main__":` entry point where appropriate.
5. Document required input, configuration, and output.
6. Add the script to the **Current Modules** table.
7. Add a short description under **Module Details**.
8. Update the technology stack only when a new dependency is introduced.

Use this template for future module documentation:

```markdown
### Module Name

**File:** `module_name.py`

**Business problem:** Explain the manual or operational issue.

**Solution:** Explain what the script automates.

**Input:** Describe files, database tables, or parameters.

**Output:** Describe generated files, logs, or actions.

**Run:**

\`\`\`bash
python module_name.py
\`\`\`
```

## Development Roadmap

Completed engineering foundations for `auto_sql.py`:

- [x] Function decomposition and responsibility separation
- [x] Custom exception architecture and exception chaining
- [x] Context-managed Oracle cursor lifecycle
- [x] Centralised application configuration
- [x] Environment-variable secret separation
- [x] `.env.example` and dependency manifest
- [x] Structured console and file logging
- [x] Timed log rotation
- [x] Log-based troubleshooting workflow

Future repository improvements:

- [ ] Add automated tests with sample or mocked Oracle data
- [ ] Add anonymised screenshots or sample report evidence
- [ ] Replace fixed UI waits with explicit Selenium conditions
- [ ] Group scripts into category folders when required

## Scope and Limitations

- The modules are currently standalone scripts rather than one deployed application.
- Oracle-connected modules require access to a configured Oracle environment.
- UI automation depends on the Oracle EBS page structure and may require selector updates.
- Sample or anonymised data should be used in this public repository.
- Scripts marked as proofs of concept should be reviewed before production use.

## Author

**FAN FAN**  
ERP / MIS Developer focused on Oracle EBS, PL/SQL, application support, and business process automation.

GitHub: [fanfan-yi](https://github.com/fanfan-yi)
