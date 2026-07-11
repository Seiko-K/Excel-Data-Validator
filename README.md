![Version](https://img.shields.io/badge/version-v0.5-blue)
![Status](https://img.shields.io/badge/status-active-success)
![License](https://img.shields.io/badge/license-MIT-green)

# Excel Data Validator

Reusable Excel VBA toolkit for data validation, quality checks, and structured validation workflows.

Built with Microsoft Excel VBA for customer, product, inventory, and research datasets.

---

## Architecture

<p align="center">
  <img src="images/architecture.svg" width="800" alt="Excel Data Validator Architecture">
</p>

---

## Features

✓ Required Field Validation

✓ Duplicate Detection

✓ Validation Status Flags

✓ Email Validation

✓ Email Domain Validation

✓ Validation Reports

✓ Summary Generation

✓ CSV Validation Export

✓ Dashboard Module

✓ Highlight Errors

✓ Validation Rules

✓ Validation Workflows

✓ Excel VBA

✓ Reusable Validation Modules

---

## Example Use Cases

### Customer Master Validation

- Missing customer names
- Missing email addresses
- Invalid email formats
- Invalid email domains
- Duplicate customer IDs
- Validation status checks
- Validation summaries
- CSV validation reports
- Dashboard-ready validation results
- Required field validation

### Product Catalog Validation

- Missing SKU
- Duplicate product codes
- Empty category fields
- Price validation
- Product quality checks

### Research Metadata Validation

- Missing DOI
- Missing title
- Duplicate records
- Metadata consistency checks

---

## Repository Structure

```text
Excel-Data-Validator/

images/
    architecture.svg

modules/
    Validation.bas
    Dashboard.bas

examples/
    sample_customer_data.csv

README.md

LICENSE
```

---

## Included Examples

### Validation.bas

Features

- Missing Name detection
- Missing Email detection
- Invalid Email detection
- Email domain validation
- Duplicate Customer ID detection
- Error highlighting
- Cleanup workflows
- Validation reports
- Summary generation
- CSV report export
- Reusable VBA modules

Location

```text
modules/Validation.bas
```

---

### Dashboard.bas

Features

- Dashboard module foundation
- Validation summary framework
- Dashboard-ready architecture
- Future UI expansion

Location

```text
modules/Dashboard.bas
```

---

## Validation Workflow

```text
ClearValidationHighlights()

↓

ValidateCustomerData()

↓

ValidateEmails()

↓

Email Domain Validation

↓

HighlightDuplicates()

↓

GenerateValidationReport()

↓

GenerateCSVReport()

↓

Dashboard Module
```

---

## Sample Dataset

**sample_customer_data.csv**

Example fields

- CustomerID
- Name
- Email

Validation scenarios

- Duplicate CustomerID
- Missing Name
- Missing Email
- Invalid Email Format
- Invalid Email Domain
- Validation Reports
- CSV Export
- Data quality checks

Example output

```csv
Row,IssueType,Column,Value
3,Duplicate,CustomerID,2
4,Missing,Name,
5,Missing,Email,
7,Invalid,Email,bobexample.com
8,Invalid,Email,john@gmail
9,Invalid,Email,john@
10,Invalid,Email,john@@example.com
11,Invalid,Email,john@example..com
```

---

## Technical Highlights

✓ Excel VBA

✓ Validation Automation

✓ Required Field Validation

✓ Email Validation

✓ Email Domain Validation

✓ Duplicate Detection

✓ CSV Export

✓ Validation Reports

✓ Dashboard Module

✓ Summary Statistics

✓ Error Highlighting

✓ Validation Status Flags

✓ Validation Workflows

✓ Reusable Macro Modules

✓ Structured Data Validation

✓ Data Quality Automation

---

## Built With

- Microsoft Excel
- VBA

---

## Roadmap

□ Dashboard UI

□ Batch Validation

□ Export Validation Results

□ Summary Export

□ ProductValidation.bas

□ ResearchValidation.bas

□ InventoryValidation.bas

□ Additional Validation Modules

---

## License

MIT © 2026 Seiko K

---

Created by Seiko-K
