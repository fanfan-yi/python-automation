# Oracle EBS Account Rule Automation

A Python-based UI automation tool designed to reduce repetitive Oracle E-Business Suite (EBS) configuration work.

This project is a sanitised portfolio version of an automation solution developed for an Oracle EBS Procurement rollout.

## Business Problem

During an Oracle EBS Procurement rollout, purchasing categories required corresponding Expense Account Rules to be configured in Oracle EBS.

The rollout involved:

- 148 purchasing categories
- 7 accounting segments per category
- 1,036 segment-level setup operations

The configuration had to be performed through the Oracle EBS web interface, where the standard data-loading approach used for other setup activities was not available for this task.

Performing the setup manually was repetitive and time-consuming, with approximately 5 hours of equivalent manual configuration effort.

## Solution

I developed an Excel-driven Python automation tool to perform the repetitive Oracle EBS UI configuration.

The workflow is:

```text
Excel Input
    ↓
Input Validation
    ↓
Oracle EBS Login
    ↓
Account Rule Creation
    ↓
Segment Selection
    ↓
Segment Value Entry
    ↓
Apply
    ↓
Continue to Next Record
```

The automation uses Selenium WebDriver to interact with the Oracle EBS web interface and pandas to process the configuration data supplied through Excel.

## Key Features

- Excel-driven configuration input
- Required-column and required-value validation
- Selenium explicit waits for more reliable UI interaction
- Support for standard segment values
- Special handling for LOV/iframe-based segment selection
- Per-record exception handling
- Recovery to a known Oracle EBS page after a failed record
- Execution summary showing successful and failed records
- Environment-based configuration for Oracle EBS URLs
- Automatic WebDriver cleanup

## Technology

- Python
- Selenium WebDriver
- pandas
- openpyxl
- pyperclip
- Oracle E-Business Suite R12.2
- Excel

## Project Structure

```text
account-rule-automation/
├── account_rule_automation.py
├── config.py
├── .env.example
└── README.md
```

## Configuration

Environment-specific values are separated from the application code.

Create a local `.env` file based on `.env.example`:

```env
EBS_LOGIN_URL=https://your-ebs-host.example.com/OA_HTML/AppsLocalLogin.jsp
EBS_TARGET_URL=https://your-ebs-host.example.com/OA_HTML/RF.jsp

AUTOMATION_INPUT_FILE=./input/account_rules.xlsx
```

Real Oracle EBS URLs, credentials, company information, and production configuration are intentionally excluded from this repository.

## Input Data

The automation expects an Excel file containing the following columns:

| Column | Purpose |
|---|---|
| `name1` | Account rule value |
| `select` | Accounting segment name |
| `unitt` | Segment value |

The program validates the input before starting the batch.

For the special `會計子目` segment, the segment value is selected through the Oracle EBS LOV interface rather than entered directly.

## Error Handling

Each Excel record is processed independently.

If a record fails:

1. The error is recorded.
2. The automation attempts to return to the Account Rule page.
3. Processing continues only if the Oracle EBS page can be recovered successfully.
4. If page recovery also fails, the batch is stopped to avoid continuing from an unknown application state.

At the end of a completed batch, the program displays an execution summary containing the total, successful, and failed record counts.

## Security and Confidentiality

This repository contains a sanitised portfolio implementation.

It does **not** contain:

- Production credentials
- Internal company URLs
- Production data
- Session information
- Company-specific configuration data
- Proprietary Oracle EBS customisations

The repository demonstrates the automation approach and engineering practices without exposing confidential production information.

## Skills Demonstrated

This project demonstrates practical experience in:

- Oracle EBS application support
- ERP configuration automation
- Python scripting
- Selenium browser automation
- Excel-based operational tooling
- Input validation
- Exception handling
- UI automation reliability
- Configuration management
- Production-support-oriented error recovery

## Background

This automation was developed as part of a broader Oracle EBS Procurement rollout.

Where Oracle-supported loading mechanisms were available, standard data-loading methods were used. Python/Selenium automation was introduced specifically for repetitive web-based configuration where the standard loading approach was not available.

The portfolio version has been redesigned to remove organisation-specific information while preserving the technical workflow and problem-solving approach.
