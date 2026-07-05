![Version](https://img.shields.io/badge/version-v0.3.1-blue)
![Status](https://img.shields.io/badge/status-active-success)
![License](https://img.shields.io/badge/license-MIT-green)

# Excel Data Validator

Excel VBA toolkit for validation, quality checks and structured data workflows.

Built for customer, product and research datasets.

---

## Features

✓ Required Field Validation

✓ Duplicate Detection

✓ Validation Status Flags

✓ Email Validation

✓ Validation Reports

✓ Summary Generation

✓ CSV Validation Export

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
- Duplicate customer IDs
- Validation status checks
- Validation summaries
- CSV validation reports
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

modules/
    Validation.bas

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

## Validation Workflow

```text
ClearValidationHighlights()

↓

ValidateCustomerData()

↓

ValidateEmails()

↓

HighlightDuplicates()

↓

GenerateValidationReport()

↓

GenerateCSVReport()
```

---

## Sample Dataset

sample_customer_data.csv

Example fields

- CustomerID
- Name
- Email

Validation scenarios

- Duplicate CustomerID
- Missing Name
- Missing Email
- Invalid Email Format
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
```

---

## Technical Highlights

✓ Excel VBA

✓ Validation Automation

✓ Duplicate Detection

✓ Email Validation

✓ CSV Export

✓ Validation Reports

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

□ Email Domain Validation

□ Batch Validation

□ Export Validation Results

□ Summary Export

□ Validation Dashboard

□ ProductValidation.bas

□ ResearchValidation.bas

□ Additional Validation Modules

---

## License

MIT © 2026 Seiko K

---

Created by Seiko-K
