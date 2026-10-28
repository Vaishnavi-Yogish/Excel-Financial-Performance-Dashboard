# Excel Financial Performance Dashboard

## Overview

An interactive Excel dashboard developed using a simulated dataset of **500 records covering 1 January 2025 to 15 May 2026**.

The project analyses revenue, profit, expenses, customer metrics, departmental performance and budget variances through PivotTables, PivotCharts and interactive slicers.

The dashboard enables performance comparisons across **17 reporting months and four regions**, with filters that update the financial summaries and charts.

## Business Problem

The analysis addresses the following questions:

- What are the total revenue, profit and expenses?
- How does financial performance vary by month, department and region?
- How does actual profit compare with budgeted profit?
- Which expense categories account for the largest share of costs?
- Which customers contribute the most revenue?
- How can these metrics be presented in a consolidated, interactive dashboard?

The objective was to turn these questions into a dashboard that allows users to review financial information without repeatedly going through the raw dataset.

## Business Solution

Developed an Excel-based reporting model using PivotTables for financial data aggregation and analysis.

The dashboard includes:

- **KPI Cards:** Revenue, actual profit, expenses, unique customers and customer satisfaction
- **Monthly Trend:** Revenue and profit across 17 separate months
- **Department Analysis:** Revenue, expenses and profit comparisons
- **Budget vs Actual:** Actual profit compared with budgeted profit by region
- **Expense Breakdown:** Distribution of expenses across categories
- **Customer Analysis:** Top 10 customers by revenue
- **Interactive Filters:** Region and Month-Year slicers connected to the dashboard's PivotTables and KPI summaries

## Business Recommendations

- **Expense Monitoring:** Review salaries and other major expense categories to understand their contribution to overall costs.
- **Departmental Performance:** Compare revenue, expenses and profitability across departments to identify areas requiring further review.
- **Budget Monitoring:** Track actual profit against budget regularly and investigate significant variances.
- **Regional Analysis:** Investigate differences in regional performance, particularly the revenue gap between West and South.

## Technical Approach

The project was built in the following sequence:

**Raw Data → KPI/Data Fields → PivotTables → PivotCharts → Dashboard → Slicers → VBA Refresh**

## Technical Skills Demonstrated

- Microsoft Excel
- PivotTables and PivotCharts
- Excel slicers and Report Connections
- KPI calculation and data aggregation
- Financial performance and budget variance analysis
- Dashboard design and visualisation
- VBA-based refresh automation

## Key Findings

The overall dataset reports:

| KPI | Value |
|---|---:|
| Total Revenue | 85.07M |
| Actual Profit | 43.35M |
| Total Expenses | 41.72M |
| Unique Customers | 243 |
| Average Customer Satisfaction | 87.47% |

### Performance Observations

- **Budget Performance:** Actual profit exceeded budgeted profit by 0.43M (0.99%), with favourable differences across all four regions.
- **Departmental Performance:** Operations recorded the highest revenue (14.85M) and profit (8.00M).
- **Regional Performance:** West contributed the highest revenue (24.00M), while South recorded the lowest (18.71M).
- **Expense Analysis:** Salaries represented approximately 26.9% of total expenses, making them the largest expense category.
- **Customer Analysis:** The top 10 customers contributed approximately 11.2% of total revenue.

The findings highlight differences in financial performance and areas requiring further investigation.

## Limitations & Assumptions

- The project uses a **simulated dataset**, not the financial records of an actual company. Therefore, the results represent the analysis of this dataset and not the performance of a real business.
- The dashboard highlights differences between departments, regions and months but does not independently establish the underlying causes.
- Departmental classifications and budget figures are based on the information provided in the dataset.
- May 2026 includes data only through 15 May and is not directly comparable with complete months.

## What I Learned

The main learning from this project was that an effective dashboard begins with identifying the business questions it needs to answer.

The project provided practical experience in:

- Structuring raw data for analysis
- Building PivotTables to answer business questions
- Creating calculated KPIs and financial comparisons
- Designing interactive charts and slicers
- Connecting multiple PivotTables to common filters
- Automating repetitive refresh tasks using VBA
- Checking whether a visual supports the conclusion being drawn from it

The project strengthened my understanding of how Excel can be used to organise financial information and support management reporting.
