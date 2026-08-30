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
- Creates a formatted Excel report
- Adds Excel filters
- Automatically formats numeric values
- Generates separate Cleaned Data and Summary sheets

## Technologies Used

- Python
- Pandas
- OpenPyXL
- Microsoft Excel
- Git
- GitHub

## Project Structure

```text
smart-excel-sales-automation/
│
├── input/
│   ├── sales_january.xlsx
│   └── sales_february.xlsx
│
├── output/
│   └── final_sales_report.xlsx
│
├── screenshots/
│   └── sales_summary.png
│
├── main.py
├── README.md
└── .gitignore