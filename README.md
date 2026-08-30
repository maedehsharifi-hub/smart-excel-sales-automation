# Smart Excel Sales Automation

A Python automation tool that reads multiple Excel sales files, validates and cleans the data, performs sales analysis, and automatically generates a formatted Excel report.

## Features

- Automatically detects multiple Excel files
- Combines sales data from multiple files
- Validates required columns
- Detects invalid numeric data
- Ignores temporary Excel files
- Removes duplicate records
- Handles missing price values
- Calculates total value for each order
- Generates sales statistics
- Creates a professional Excel report
- Adds Excel filters and formatting

## Technologies

- Python
- Pandas
- OpenPyXL
- Excel

## Input Format

Excel files should contain the following columns:

- Name
- Email
- Product
- Quantity
- Price

## Output

The program generates:

output/final_sales_report.xlsx

The report contains two sheets:

### Cleaned Data

Contains the combined and cleaned sales records.

### Summary

Contains:

- Original rows
- Duplicates removed
- Missing prices removed
- Total orders
- Total sales
- Average order value
- Highest order value

## How to Run

Install dependencies:

```bash
pip install pandas openpyxl