Vikings is an Android mobile phone and accessories wholesale distribution company in North Bangalore that distributes to hundreds of retailers within the region. Our mission is to provide high-quality products and exceptional service to our partners, ensuring they have the best offerings for their customers. With a strong focus on growth and innovation, Vikings is committed to supporting retailers in achieving their business goals.


# Viking Analytics

Viking Analytics is an intuitive and robust reporting system that generates daily reports on Growth, Credit, COGS, Sales, Inventory Management, Price Lists, Zero Stock Orders (ZSO), and RA Norms, providing valuable insights for informed decision-making. It provides crucial insights that empower businesses to make informed decisions, track performance, and manage risks effectively. By leveraging comprehensive data analysis, Viking Reports helps organizations optimize operations, enhance financial management, and drive sustainable growth. The reports are produced in Excel format, facilitating easy distribution.

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

6. **Zero Stock Order (ZSO) Management**:
   - Identify ZSO situations where specific models have zero inventory
   - Enable proactive inventory management to prevent stockouts
   - Support data-driven decisions for inventory replenishment and risk mitigation

7. **RA Norms**:
   - RA Norms are used to determine the minimum inventory levels for retailers based on sales trends and required stock replenishment.
   - Helps optimize inventory management by ensuring each retailer keeps a minimum stock of models that are in demand.
   - Supports proactive decision-making by identifying retailers that need to restock and prevent potential stockouts.   

## Key Features

- **Growth Report**: Analyzes SO and ST data, calculates growth percentages, and generates color-coded Excel reports for easy interpretation.
- **Credit Report**: Processes bill data, categorizes by age, and creates separate reports for each TSE and missing TSE data.
- **COGS Report**: Calculates the total inventory cost for each retailer, assesses potential financial exposure in case of flight risk, and provides insights for inventory management.
- **Sales Report**: Fetches sales data from Excel files for each retailer, computes month-to-date sales performance, and generates detailed reports categorized by product types.
- **PriceList Report**: Generates a flat price list of SKUs for the current month, fetching price data from the zonal distributor and inventory data for material codes. Outputs a structured Excel report containing SKU Type, Model, Color, Variant, NLC (Net Landing Cost), MOP (Market Operating Price), MRP (Maximum Retail Price), and Material Code.
- **ZSO Report**: Identifies Zero Stock Order (ZSO) situations where specific models have zero inventory, combines sales and inventory data to determine ZSO status for each dealer, and generates an Excel report highlighting ZSO models and their respective dealers.
- **RA Norms Report**: Identifies which retailers need to restock based on RA norms, considering factors like current inventory levels, sell-through data, and growth patterns. Generates a report that helps businesses optimize inventory replenishment.
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
   - Categorizes bills by age (0-7 days, 8-14 days, 15-20 days, 21-30 days, >30 days)
   - Compute the amount collected from retailers previous day
   - Aggregates credit data by retailer and TSE (Territory Sales Executive)
   - Generates separate Excel reports for each TSE and a report for missing TSE data
3. COGS Report Generation
   - Calculates the total inventory cost for each retailer (store)
   - Helps assess potential financial impact in case of flight risk by computing inventory shortfall
   - Provides insights for inventory management and risk assessment
4. Sales Report Generation
   - Fetches sales data from Excel files for each retailer
   - Computes month-to-date sales performance
   - Generates reports categorized by product types (e.g., SMART PHONES, ACCESSORIES)
   - Outputs detailed sales reports with total sales values and quantities for each retailer
   - Supports TSE-specific reporting for targeted follow-ups
5. PriceList Report Generation
   - Generates a flat price list of SKUs for the current month.
   - Fetches price data from the zonal distributor and inventory data for material codes.
   - Outputs a structured Excel report containing:
     - SKU Type
     - Model
     - Color
     - Variant
     - NLC (Net Landing Cost)
     - MOP (Market Operating Price)
     - MRP (Maximum Retail Price)
     - Material Code
6. ZSO Report Generation
   - Identifies Zero Stock Out (ZSO) situations where specific models have zero inventory.
   - Combines sales data (Sell Out of LMTD, MTD) and inventory data to determine ZSO status for each dealer.
   - Generates an Excel report highlighting ZSO models and their respective dealers.
   - Supports proactive inventory management and risk mitigation strategies.     
7. RA Norms Report Generation
   - Analyzes sales trends and inventory data to identify which retailers need to restock specific mobile models based on RA norms.
   - Computes the minimum stock level required for each retailer, ensuring that they have adequate inventory levels to meet customer demand.
   - Generates a report that helps prioritize inventory replenishment and avoid stockouts.
   - Provides actionable insights to ensure retailers adhere to RA norms.   

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
   - `data/MTD-SO.xlsx`
   - `data/LMTD-SO.xlsx`
   - `data/MTD-ST.xlsx`
   - `data/LMTD-ST.xlsx`
   - `data/Retailer Metadata.xlsx`

2. Run the growth report generator:
   ```
   cd growth_report
   go run main.go
   ```

3. The generated report will be saved in a new directory named `growth_reports_YYYY-MM-DD`.

### Credit Report

1. Ensure the following Excel files are present in the `data` directory:
   - `data/Bills.xlsx`
   - `data/Retailer Metadata.xlsx`

2. Run the credit report generator:
   ```
   cd credit_report
   go run main.go
   ```

3. The generated reports will be saved in a new directory named `credit_reports_YYYY-MM-DD`.

### COGS Report

1. Ensure the following Excel file is present in the `data` directory:
   - `data/DealerInventory.xlsx` (containing current inventory data for all retailers)
   - `data/ProductPriceList.xlsx`
   - `data/Retailer Metadata.xlsx`
   - `generated credit report of each tse`


2. Run the COGS report generator:
   ```
   cd cogs_report
   go run main.go
   ```

3. The generated report will be saved in a new directory named `cogs_reports_YYYY-MM-DD`.

### Sales Report

1. Ensure the following Excel file is present in the `data` directory:
   - `data/Sales.xlsx` (containing sales data for each retailer)

2. Run the sales report generator:
   ```
   cd sales_report
   go run main.go
   ```

3. The generated sales report will be saved in a new directory named `sales_reports_YYYY-MM-DD`.

### PriceList Report

1. Ensure the following Excel files are present in the `data` directory:
   - `data/ProductPriceList.xlsx` (containing price data from the zonal distributor)
   - `data/DealerInventory.xlsx` (containing current inventory data for all retailers)

2. Run the price list report generator:
   ```
   cd pricelist
   go run main.go
   ```

3. The generated report will be saved in a new directory named `price_list_reports_YYYY-MM-DD`.

### ZSO Report

1. Ensure the following Excel files are present in the `data` directory:
   - `data/L2M-SO.xlsx`
   - `data/DealerInventory.xlsx` (containing current inventory data for all retailers)

2. Run the ZSO report generator:
   ```
   cd zso_report
   go run main.go
   ```

3. The generated report will be saved in a new directory named `zso_reports_YYYY-MM-DD`.

### RA Norms Report
1. Ensure the following Excel files are present in the data directory:
   - `data/Sales.xlsx` (containing sales data for each retailer)
   - `data/DealerInventory.xlsx` (containing current inventory data for all retailers)
   - `data/Retailer Metadata.xlsx`

2. Run the RA norms report generator:
   ```
   cd ranorms
   go run main.go
   ```
3. The generated RA Norms report will be saved in a new directory named `ranorms_reports_YYYY-MM-DD`.

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