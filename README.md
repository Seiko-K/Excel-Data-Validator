![Version](https://img.shields.io/badge/version-v0.8-blue)
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

✓ Customer Data Validation

✓ Product Data Validation

✓ Research Metadata Validation

✓ Inventory Data Validation

✓ Duplicate Detection

✓ Validation Status Flags

✓ Email Validation

✓ Email Domain Validation

✓ Validation Reports

✓ Summary Generation

✓ CSV Validation Export

✓ Validation Dashboard

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
- Dashboard-ready validation results
- CSV validation reports
- Required field validation

### Product Catalog Validation

- Missing SKU
- Missing product names
- Duplicate product codes
- Invalid prices
- Negative prices
- Product quality checks

### Research Metadata Validation

- Missing DOI
- Missing title
- Invalid publication year
- Future publication year
- Metadata consistency checks
- Research dataset validation

### Inventory Validation

- Missing SKU
- Missing product names
- Negative stock quantity
- Invalid stock quantity
- Negative reorder level
- Inventory data validation

---

## Repository Structure

```text
Excel-Data-Validator/

images/
    architecture.svg

modules/
    Validation.bas
    Dashboard.bas
    ProductValidation.bas
    ResearchValidation.bas
    InventoryValidation.bas

examples/
    sample_customer_data.csv
    sample_product_data.csv
    sample_research_data.csv
    sample_inventory_data.csv

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

- Validation dashboard
- Validation summary display
- Dashboard metrics
- Dashboard foundation
- Future dashboard UI support

Location

```text
modules/Dashboard.bas
```

---

### ProductValidation.bas

Features

- Missing SKU detection
- Missing Product Name detection
- Invalid Price detection
- Negative Price detection
- Product validation highlighting
- Reusable product validation module

Location

```text
modules/ProductValidation.bas
```

---

### ResearchValidation.bas

Features

- Missing DOI detection
- Missing Title detection
- Invalid Publication Year detection
- Future Publication Year detection
- Research metadata highlighting
- Reusable research validation module

Location

```text
modules/ResearchValidation.bas
```

---

### InventoryValidation.bas

Features

- Missing SKU detection
- Missing Product Name detection
- Invalid Stock Quantity detection
- Negative Stock Quantity detection
- Invalid Reorder Level detection
- Negative Reorder Level detection
- Inventory validation highlighting
- Reusable inventory validation module

Location

```text
modules/InventoryValidation.bas
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

ValidateProductData()

↓

ValidateResearchData()

↓

ValidateInventoryData()

↓

HighlightDuplicates()

↓

GenerateValidationReport()

↓

GenerateCSVReport()

↓

ShowValidationDashboard()
```

---

## Sample Datasets

### sample_customer_data.csv

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
- Dashboard summary
- CSV Export

---

### sample_product_data.csv

Example fields

- SKU
- ProductName
- Price

Validation scenarios

- Missing SKU
- Missing Product Name
- Negative Price
- Invalid Price
- Product validation
- Dashboard summary

---

### sample_research_data.csv

Example fields

- DOI
- Title
- PublicationYear

Validation scenarios

- Missing DOI
- Missing Title
- Invalid Publication Year
- Future Publication Year
- Research metadata validation
- Dashboard summary

---

### sample_inventory_data.csv

Example fields

- SKU
- ProductName
- StockQuantity
- ReorderLevel

Validation scenarios

- Missing SKU
- Missing Product Name
- Negative Stock Quantity
- Invalid Stock Quantity
- Negative Reorder Level
- Invalid Reorder Level
- Inventory validation
- Dashboard summary

Example output

```csv
Row,IssueType,Column,Value
3,Missing,SKU,
4,Missing,ProductName,
5,Invalid,StockQuantity,-3
6,Invalid,StockQuantity,invalid
7,Invalid,ReorderLevel,-1
8,Invalid,ReorderLevel,invalid
```

---

## Technical Highlights

✓ Excel VBA

✓ Validation Automation

✓ Customer Data Validation

✓ Product Data Validation

✓ Research Metadata Validation

✓ Inventory Data Validation

✓ Email Validation

✓ Email Domain Validation

✓ Duplicate Detection

✓ Validation Dashboard

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

□ Dashboard Charts

□ Dashboard UI

□ Batch Validation

□ Export Validation Results

□ Summary Export

□ Additional Validation Modules

---

## License

MIT © 2026 Seiko K

---

Created by Seiko-K
