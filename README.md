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
| [`auto_sql.py`](./auto_sql.py) | Data quality | Detects `PO_DISTRIBUTIONS_ALL` records with missing requisition and delivery information | Date-stamped exception report | Oracle environment required |
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

### 3. PO Distribution Data-Quality Check

**File:** `auto_sql.py`

Queries `PO_DISTRIBUTIONS_ALL` for records where requisition and delivery-related fields are missing. When exceptions exist, the script exports them to a date-stamped Excel workbook for investigation.

This supports a practical application-support workflow:

1. Detect abnormal records.
2. Export evidence for analysis.
3. Investigate affected purchase orders.
4. Correct the process or data issue.

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
| Database connectivity | cx_Oracle |
| UI automation | Selenium, PyAutoGUI, Pyperclip |

## Repository Structure

```text
python-automation/
├── README.md
├── po_headers_report.py
├── PO_Count.py
├── auto_sql.py
├── creat_office.py
└── output/                   # Generated reports; created when required
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

Some scripts require local configuration before they can run:

- Oracle Instant Client path;
- database host, port, and service name;
- database username and password;
- ChromeDriver and browser compatibility;
- Oracle EBS URLs;
- local input and output paths.

Never commit real credentials, internal URLs, session parameters, or production data. Use environment variables or a secure secrets store for sensitive configuration.

Example future environment variables:

```dotenv
ORACLE_USER=your_username
ORACLE_PASSWORD=your_password
ORACLE_DSN=host:port/service_name
ORACLE_CLIENT_PATH=your_oracle_client_path
```

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

- [ ] Move sensitive configuration to environment variables
- [ ] Add `.env.example`
- [ ] Add `requirements.txt`
- [ ] Add shared database connection utilities
- [ ] Add structured logging
- [ ] Improve connection cleanup and exception handling
- [ ] Replace fixed UI waits with explicit Selenium conditions
- [ ] Add automated tests with sample data
- [ ] Add screenshots or anonymised sample reports
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
