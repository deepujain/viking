# Viking Reports

This project contains two main components for generating reports: Growth Report and Credit Report.

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Dependencies](#dependencies)

## Features

1. Growth Report Generation
   - Analyzes sell-out (SO) and sell-through (ST) data
   - Compares current month-to-date (MTD) with last month-to-date (LMTD) data
   - Calculates growth percentages
   - Generates Excel report with color-coded growth indicators

2. Credit Report Generation
   - Processes bill data for retailers
   - Categorizes bills by age (0-7 days, 8-14 days, 15-21 days, 22-30 days, >30 days)
   - Aggregates credit data by retailer and TSE (Territory Sales Executive)
   - Generates separate Excel reports for each TSE and a report for missing TSE data

## Prerequisites

- Go 1.x or higher
- Excel files with required data (see Usage section for file names)

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/your-username/viking-reports.git
   cd viking-reports
   ```

2. Install dependencies:
   ```
   go mod tidy
   ```

## Usage

### Growth Report

1. Ensure the following Excel files are present in the `data` directory:
   - `Dealer Information.xlsx`
   - `MTD-SO.xlsx`
   - `LMTD-SO.xlsx`
   - `MTD-ST.xlsx`
   - `LMTD-ST.xlsx`
   - `VIKING'S - DEALER Credit Period LIST.xlsx`

2. Run the growth report generator:
   ```
   cd growth_report
   go run main.go
   ```

3. The generated report will be saved in a new directory named `daily_growth_reports_YYYY-MM-DD`.

### Credit Report

1. Ensure the following Excel files are present in the `data` directory:
   - `Bills.xlsx`
   - `VIKING'S - DEALER Credit Period LIST.xlsx`

2. Run the credit report generator:
   ```
   cd credit_report
   go run main.go
   ```

3. The generated reports will be saved in a new directory named `daily_credit_reports_YYYY-MM-DD`.

## Dependencies

- [github.com/xuri/excelize/v2](https://github.com/xuri/excelize): Used for reading and writing Excel files.