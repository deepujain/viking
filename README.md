# Viking Reports

Viking Reports is a robust reporting system designed to generate daily Growth and Credit reports, providing crucial insights for business decision-making and performance tracking.

## Table of Contents

- [Business Objectives](#business-objectives)
- [Key Features](#key-features)
- [Reports](#reports)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Dependencies](#dependencies)


## Business Objectives

Our daily reporting system addresses key business needs:

1. **Growth Tracking**: 
   - Monitor daily sales performance across retailers
   - Compare current month-to-date (MTD) with last month-to-date (LMTD) data
   - Identify growth trends in sell-out (SO) and sell-through (ST) metrics

2. **Credit Management**:
   - Track outstanding bills by retailer
   - Categorize bills by age for better debt management
   - Provide Territory Sales Executive (TSE) specific reports for targeted follow-ups

3. **Performance Analysis**:
   - Highlight top-performing and underperforming retailers
   - Enable data-driven decisions for inventory management and sales strategies

4. **Risk Assessment**:
   - Identify retailers with aging credit for proactive risk mitigation
   - Support financial planning and cash flow management

## Key Features

- **Growth Report**: Analyzes SO and ST data, calculates growth percentages, and generates color-coded Excel reports for easy interpretation.
- **Credit Report**: Processes bill data, categorizes by age, and creates separate reports for each TSE and missing TSE data.
- **Automated Daily Generation**: Reports are automatically generated and saved with date-stamped directories for easy tracking and comparison.

By providing these daily reports, Viking Reports empowers businesses to make informed decisions, optimize operations, and drive growth while managing financial risks effectively.

## Reports

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