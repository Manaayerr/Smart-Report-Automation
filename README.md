
# Smart Report Automation

An automated Python tool designed to read raw transaction data from a CSV file, perform key business calculations, and generate a polished, highly formatted multi-tab Excel dashboard with dynamic charts and conditional formatting.

## Features

* **Automated Data Processing**: Uses `pandas` to clean data, calculate monthly aggregates, and analyze product categories.
* **Dynamic Excel Dashboard**: Automatically generates an interactive executive summary sheet featuring KPI cards.
* **Data Visualization**: Embeds automated, responsive charts comparing monthly Revenue and Profit performance.
* **Professional Styling**: Features custom-styled tables, zebra striping, consistent font scaling (Segoe UI), auto-fitted columns, and double-line totals borders.
* **Conditional Formatting**: Applies soft, professional visual indicators to highlight product lines exceeding a 40% profit margin.
* **Interactive Raw Ledger**: Keeps a separate sheet with frozen headers for quick and seamless raw transactional audits.

## Project Structure

Your project directory should be structured as follows:

```text
Smart-Report-Automation/
│
├── raw_sales_data.csv          # Raw transaction input dataset
├── report_generator.py         # Main Python automation script
├── sales_performance_report.xlsx # Generated Excel dashboard output
└── README.md                   # Project documentation

```

## Getting Started

Follow these instructions to set up and run the automated report generator on your local machine.

### Prerequisites

Make sure you have Python installed. You will also need to install the required external libraries:

```bash
pip install pandas openpyxl numpy

```

### Usage

1. Place your input transaction data inside the project folder and name it `raw_sales_data.csv`.
2. Run the main automation script:

```bash
python report_generator.py

```

3. Once the execution completes, you will find the fully styled spreadsheet `sales_performance_report.xlsx` successfully generated in your project folder.

## Technologies Used

* **Python 3**
* **Pandas**: Data manipulation and analytics.
* **OpenPyXL**: Excel workbook layout structure, styles, charts, and formatting.

```

```
