# Bank Loan Analysis (Excel)

A bank loan risk analysis built in Excel using **PivotTables, PivotCharts and Slicers**. The workbook analyses **350 loans** across 7 branches to understand loan exposure, overdue loans and risk ratings.

> **Note:** This is a practice dataset created for learning. It is not real customer data.

## Project Objective

To answer questions a bank risk team would ask:

- How much has the bank lent, and where?
- Which branches have the most overdue loans?
- Does the risk rating match actual repayment behaviour?
- How has lending changed over time?

## Dataset

| Column | Description |
|---|---|
| Loan_ID, Customer_ID, Customer_Name | Identifiers |
| Branch | Colombo, Galle, Jaffna, Kandy, Kurunegala, Matara, Negombo |
| Loan_Type | Business, Education, Home, Personal, Vehicle |
| Loan_Amount_LKR | Loan amount in Sri Lankan rupees |
| Interest_Rate_% | Interest rate |
| Loan_Date | January 2023 to September 2026 |
| Status | Active, Overdue, Paid |
| Risk_Rating | Low, Medium, High |
| Loan_Officer | Officer handling the loan |
| Monthly_Income_LKR | Customer monthly income |

**Data quality checks:** 350 rows, no missing values, no duplicate Loan_IDs, no extra spaces in the main columns.

## What is inside the workbook

| Sheet | Content |
|---|---|
| Loan_Data | The raw data (350 rows) |
| Pivot_Tables | 9 PivotTables |
| Dashboard | 7 PivotCharts with Branch and Risk_Rating slicers |

## Key Findings

- The total loan portfolio is **LKR 928.9 million** across 350 loans.
- **107 loans (30.6%)** are overdue, worth about **26%** of the total loan amount.
- **Matara** has the highest overdue amount (LKR 48.8M). **Negombo** has the lowest (LKR 19.0M).
- **High-risk loans are about 51% overdue**, compared with about 20% for Low and Medium risk loans.
- The average interest rate is almost the same for every risk rating (about 9.4% to 9.6%), so higher-risk loans are **not priced higher**.
- **Home loans** are the largest portfolio at LKR 413.2M (44.5% of the total).
- Loan amount by year falls from LKR 298.7M (2023) to LKR 159.5M (2026). **2026 only includes data up to September**, so it is not a full-year comparison.



## Files in this repository

| File | Description |
|---|---|
| `Bank_Loan_Analysis.xlsx` | The full workbook (download to view PivotTables, charts and slicers) |
| `loan_data.csv` | The raw data in CSV format |
| `images/` | Screenshots of the analysis |

## Excel skills demonstrated

- PivotTables, PivotCharts and Slicers
- Lookup functions: VLOOKUP, XLOOKUP
- Error handling: IFERROR, IFNA
- Conditional calculations: COUNTIFS, SUMIFS, AVERAGEIFS

## How to use

1. Download `Bank_Loan_Analysis.xlsx`.
2. Open it in Microsoft Excel.
3. Use the Branch and Risk_Rating slicers on the Dashboard sheet to filter the charts.
4. If you change the data in Loan_Data, go to **Data > Refresh All** to update the PivotTables and charts.
