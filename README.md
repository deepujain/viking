# Viking Reports

Viking Reports is a comprehensive reporting system designed to generate daily Growth, Credit, and COGS reports, providing crucial insights for business decision-making, performance tracking, and risk management.

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
     - Sell-through (ST): Sale of Viking's mobiles and accessories to its retailers
     - Sell-out (SO): Sale of those mobiles and accessories by retailers to customers


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

5. **Inventory Cost Management**:
   - Track total inventory cost for each retailer
   - Assess potential financial exposure in case of retailer flight risk
   - Support inventory optimization and risk mitigation strategies

## Key Features

- **Growth Report**: Analyzes SO and ST data, calculates growth percentages, and generates color-coded Excel reports for easy interpretation.
- **Credit Report**: Processes bill data, categorizes by age, and creates separate reports for each TSE and missing TSE data.
- **Automated Daily Generation**: Reports are automatically generated and saved with date-stamped directories for easy tracking and comparison.

By providing these daily reports, Viking Reports empowers businesses to make informed decisions, optimize operations, and drive growth while managing financial risks effectively.

## Reports

1. Growth Report Generation
   - Analyzes sell-out (SO) and sell-through (ST) data
     - Sell-through (ST): Viking's sales to its retailers
     - Sell-out (SO): Retailers' sales to end customers
   - Compares current month-to-date (MTD) with last month-to-date (LMTD) data
   - Calculates growth percentages
   - Generates Excel report with color-coded growth indicators

2. Credit Report Generation
   - Processes bill data for retailers
   - Categorizes bills by age (0-7 days, 8-14 days, 15-21 days, 22-30 days, >30 days)
   - Aggregates credit data by retailer and TSE (Territory Sales Executive)
   - Generates separate Excel reports for each TSE and a report for missing TSE data
3. COGS Report Generation
   - Calculates the total inventory cost for each retailer (store)
   - Helps assess potential financial impact in case of flight risk
   - Provides insights for inventory management and risk assessment

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
   - `data/common/MTD-SO.xlsx`
   - `data/common/LMTD-SO.xlsx`
   - `data/common/MTD-ST.xlsx`
   - `data/common/LMTD-ST.xlsx`
   - `data/common/VIKING'S - DEALER Credit Period LIST.xlsx`

2. Run the growth report generator:
   ```
   cd growth_report
   go run main.go
   ```

3. The generated report will be saved in a new directory named `daily_growth_reports_YYYY-MM-DD`.

### Credit Report

1. Ensure the following Excel files are present in the `data` directory:
   - `data/credit_report/Bills.xlsx`
   - `data/common/VIKING'S - DEALER Credit Period LIST.xlsx`

2. Run the credit report generator:
   ```
   cd credit_report
   go run main.go
   ```

3. The generated reports will be saved in a new directory named `daily_credit_reports_YYYY-MM-DD`.

### COGS Report

1. Ensure the following Excel file is present in the `data` directory:
   - `data/cogs_report/DealerInventory.xlsx` (containing current inventory data for all retailers)
   - `data/common/ProductPriceList.xlsx`
   - `data/common/VIKING'S - DEALER Credit Period LIST.xlsx`
   - `generated credit report of each tse`


2. Run the COGS report generator:
   ```
   cd cogs_report
   go run main.go
   ```

3. The generated report will be saved in a new directory named `daily_cogs_reports_YYYY-MM-DD`.

## Dependencies

- [github.com/xuri/excelize/v2](https://github.com/xuri/excelize): Used for reading and writing Excel files.

## License

This project is proprietary software owned by Viking's. All rights reserved.

### Restrictions

1. **Commercial Use**: Any commercial use, reproduction, or distribution of this codebase is strictly prohibited without explicit written permission from Viking's.

2. **Proprietary Information**: All reports generated by this system are proprietary information of Viking's and may not be shared, distributed, or used outside of Viking's without express authorization.

3. **Modifications**: No modifications or derivative works based on this codebase are allowed without prior written consent from Viking's.

4. **Confidentiality**: Users of this software agree to maintain the confidentiality of the codebase and any associated documentation.

For any licensing inquiries or permissions, please contact Viking's.

© 2024 Viking's. All Rights Reserved.