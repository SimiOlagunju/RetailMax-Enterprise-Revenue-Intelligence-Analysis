# RetailMax-Enterprise-Revenue-Intelligence
Excel + Power BI Project | Sales Performance, Customer Intelligence & Revenue Optimisation

### Table of Contents
* [Project Overview](#project-overview)
* [Tools Used](#tools-used)
* [Dataset Overview](#dataset-overview)
* [Data Cleaning](#data-cleaning)
* [Exploratory Analysis](#exploratory-analysis)
* [Power BI Dashboard](#power-bi-dashboard)
* [Key Metrics](#key-metrics)
* [Key Insights](#key-insights)
* [Recommendations](#recommendations)
* [Dataset Source](#dataset-source)

### Project Overview
This project analyses retail transaction data for RetailMax Group Ltd to uncover sales performance trends, customer behaviour patterns, channel profitability, discount impact and revenue leakage caused by returns and cancellations.
> The goal is to:
> - Analyse total revenue, gross profit and net profit across products, regions and sales channels
> - Identify revenue leakage from returned and cancelled orders
> - Understand customer segmentation, loyalty patterns and lifetime value
> - Evaluate the impact of discounting on profitability across categories and channels
> - Support data-driven commercial decisions for retail leadership and strategy teams

### Tools Used
- Microsoft Excel – Data cleaning, validation and preparation
- Power BI – Data modelling, DAX measure development and interactive 5-page dashboard

### Dataset Overview
The dataset contains retail transaction and customer data including:
> - Sales transaction details (channel, order status, discount band)
> - Product information (category, brand)
> - Customer profiles (segment, age band, gender, region, city, loyalty status)
> - Store information (store name, region, total stores)
> - Date dimension (year, month)
> - Revenue, profit and discount metrics

- Total Tables: 5 (Sales Transactions, Products, Customers, Stores, Date)

Sample Overview

Sales Transactions Data
Transaction_ID, Customer_ID, Product_ID, Store_ID, Sales_Channel, Order_Status, Discount_Band, Amount, Date

Products Data
Product_ID, Category, Brand

Customers Data
Customer_ID, Customer_Segment, Age_Band, Gender, City, Region, Loyalty_Status

Stores Data
Store_ID, Store_Name, Region

Sales Transactions Sample Preview
| Transaction_ID | Customer_ID | Sales_Channel | Order_Status | Discount_Band | Amount |
|----------------|-------------|---------------|--------------|---------------|--------|
| TXN00001 | C001 | Online | Completed | Medium | 249.99 |
| TXN00002 | C002 | In-Store | Returned | None | 89.50 |
| TXN00003 | C003 | Wholesale | Cancelled | High | 1,200.00 |

Customers Sample Preview
| Customer_ID | Customer_Segment | Age_Band | Gender | Region | Loyalty_Status |
|-------------|-----------------|----------|--------|--------|----------------|
| C001 | Premium | 25-34 | Female | North | Gold |
| C002 | Standard | 35-49 | Male | South | Silver |

### Data Cleaning
Key cleaning and preparation steps performed in Power Query within Power BI:
- Removed duplicate transaction records across all 5 tables
- Handled null and missing values across Sales Transactions, Customers, Products and Stores tables
- Standardised sales channel, order status and discount band values for consistency
- Converted date fields into proper datetime formats
- Verified referential integrity between Sales Transactions, Products, Customers and Stores tables

Additional transformations performed in Power BI using DAX:
- Created a Date dimension table (DateTable) for time intelligence analysis covering 2023 to 2025
- Created a Discount Band calculated column in the Sales Transactions table using DAX IF logic grouping discounts into: No Discount, 0-10%, 10-20%, 20-30%, 30%+
- Created an Age Band calculated column in the Customers table grouping ages into: Under 25, 25-34, 35-44, 45-54, 55-64, 65-74, 75+
- Created 7 DAX measure tables to organise all calculations by business area:
> - Revenue Measures (_MyRevMeasures): Total Revenue, Gross Profit, Net Profit
> - Sales Measures (_MysalesMeasures): Total Sales, Gross Profit, Net Profit, Net Profit Margin %, Gross Potential Revenue
> - Transaction Measures (_MyTransMeasures): Total Orders, Average Transaction Value
> - Returns & Cancellations Measures (_MyRetCanMeasures): Returned Revenue, Cancelled Revenue, Revenue Leakage, Revenue Lost R&C, Sales Leakage
> - Customer Measures (_MyCustMeasures): Total Customers, Total Active Customers, Total Inactive Customers, Avg CLV, Repeat Purchase Rate %
> - Channel Measures (_MyChannelMeasures): Online Revenue - Gross Potential, Actual Online Revenue
> - Discount Measures (_MyDiscountMeasures): Avg Discount %, Total Discount Value, Discount Revenue Impact %, Profit Before Discount, No Discount Revenue, Profit Margin by Band

Key DAX measures created:
- Gross Potential Revenue — sum of all revenue regardless of order status
- Revenue Leakage — Gross Potential Revenue minus Total Revenue
- Revenue Leakage % — Revenue Leakage divided by Gross Potential Revenue
- Online Revenue — CALCULATE filtering Sales Channel to Online only
- Actual Online Revenue — CALCULATE filtering Online completed orders only
- Returned Revenue — SUMX filtering Order Status to Returned
- Cancelled Revenue — SUMX filtering Order Status to Cancelled
- Completed Order Count, Returned Order Count, Cancelled Order Count
- No Discount Revenue — CALCULATE filtering Discount Percent to zero
- Profit Margin by Band — Net Profit divided by Total Revenue
- Inactive Customers — Total Customers minus Total Active Customers

### Exploratory Analysis
All exploratory analysis was performed using DAX measures and Power BI visuals. The following summarises the key business questions explored and the DAX logic used to answer them.

1. Overall revenue and profitability summary
```dax
-- Gross Potential Revenue (all orders regardless of status)
Gross Potential Revenue =
SUMX(
    FILTER('Sales_Transactions_Table','Sales_Transactions_Table'[Revenue] >= 0),
    'Sales_Transactions_Table'[Revenue]
)

-- Total Revenue (completed orders only)
Total Revenue =
CALCULATE(
    SUM('Sales_Transactions_Table'[Revenue]),
    'Sales_Transactions_Table'[Order_Status] = "Completed"
)

-- Revenue Leakage
Revenue Leakage = [Gross Potential Revenue] - [Total Revenue]

-- Revenue Leakage %
Revenue Leakage % = DIVIDE([Revenue Leakage], [Gross Potential Revenue], 0)
```

2. Revenue and profit by sales channel
```dax
-- Online Revenue (gross potential)
Online Revenue =
CALCULATE(
    [Gross Potential Revenue],
    'Sales_Transactions_Table'[Sales_Channel] = "Online"
)

-- Actual Online Revenue (completed only)
Actual Online Revenue =
CALCULATE(
    [Total Revenue],
    'Sales_Transactions_Table'[Sales_Channel] = "Online"
)

-- Net Profit Margin %
Net Profit Margin % = DIVIDE([Net Profit], [Total Revenue], 0)
```

3. Revenue lost to returns and cancellations
```dax
-- Returned Revenue
Returned Revenue =
SUMX(
    FILTER(
        'Sales_Transactions_Table',
        'Sales_Transactions_Table'[Order_Status] = "Returned"
        && 'Sales_Transactions_Table'[Revenue] >= 0
    ),
    'Sales_Transactions_Table'[Revenue]
)

-- Cancelled Revenue
Cancelled Revenue =
SUMX(
    FILTER(
        'Sales_Transactions_Table',
        'Sales_Transactions_Table'[Order_Status] = "Cancelled"
        && 'Sales_Transactions_Table'[Revenue] >= 0
    ),
    'Sales_Transactions_Table'[Revenue]
)
```

4. Order volume by status
```dax
-- Completed Order Count
Completed Order Count =
CALCULATE(
    COUNTROWS('Sales_Transactions_Table'),
    'Sales_Transactions_Table'[Order_Status] = "Completed"
)

-- Returned Order Count
Returned Order Count =
CALCULATE(
    COUNTROWS('Sales_Transactions_Table'),
    'Sales_Transactions_Table'[Order_Status] = "Returned"
)

-- Cancelled Order Count
Cancelled Order Count =
CALCULATE(
    COUNTROWS('Sales_Transactions_Table'),
    'Sales_Transactions_Table'[Order_Status] = "Cancelled"
)
```

5. Customer segmentation and lifetime value
```dax
-- Total Active Customers
Total Active Customers =
CALCULATE(
    DISTINCTCOUNT('Sales_Transactions_Table'[Customer_ID]),
    'Sales_Transactions_Table'[Order_Status] = "Completed"
)

-- Inactive Customers
Inactive Customers =
CALCULATE(COUNTROWS('Customers_Table')) - [Total Active Customers]

-- Average Customer Lifetime Value
Avg CLV =
DIVIDE([Total Revenue], [Total Active Customers], 0)

-- Repeat Purchase Rate %
Repeat Purchase Rate % =
DIVIDE(
    COUNTROWS(FILTER(
        SUMMARIZE('Sales_Transactions_Table', 'Sales_Transactions_Table'[Customer_ID],
        "OrderCount", COUNTROWS('Sales_Transactions_Table')),
        [OrderCount] > 1
    )),
    [Total Active Customers],
    0
) * 100
```

6. Discount impact on profitability
```dax
-- Discount Band (Calculated Column in Sales table)
Discount Band =
IF('Sales_Transactions_Table'[Discount_Percent] = 0, "No Discount",
IF('Sales_Transactions_Table'[Discount_Percent] <= 0.1, "0-10%",
IF('Sales_Transactions_Table'[Discount_Percent] <= 0.2, "10-20%",
IF('Sales_Transactions_Table'[Discount_Percent] <= 0.3, "20-30%", "30%+"))))

-- No Discount Revenue
No Discount Revenue =
CALCULATE(
    [Total Revenue],
    'Sales_Transactions_Table'[Discount_Percent] = 0
)

-- Profit Margin by Discount Band
Profit Margin by Band = DIVIDE([Net Profit], [Total Revenue], 0)
```

7. Age band segmentation
```dax
-- Age Band (Calculated Column in Customers table)
Age Band =
IF('Customers_Table'[Age] <= 24, "Under 25",
IF('Customers_Table'[Age] <= 34, "25-34",
IF('Customers_Table'[Age] <= 44, "35-44",
IF('Customers_Table'[Age] <= 54, "45-54",
IF('Customers_Table'[Age] <= 64, "55-64",
IF('Customers_Table'[Age] <= 74, "65-74", "75+"))))))
```

8. Gross profit analysis
```dax
-- Gross Profit
Gross Profit = [Total Revenue] - [Total Cost]

-- Average Transaction Value
Average Transaction Value =
DIVIDE([Total Revenue], [Completed Order Count], 0)
```

9. Sales leakage analysis
```dax
-- Sales Leakage
Sales Leakage = [Gross Potential Revenue] - [Total Revenue]

-- Discount Revenue Impact %
Discount Revenue Impact % =
DIVIDE([Total Discount Value], [Gross Potential Revenue], 0) * 100

-- Profit Before Discount
Profit Before Discount = [Gross Profit] + [Total Discount Value]
```

10. Store and regional performance
```dax
-- Total Stores
Total Stores = COUNTROWS('Store_Table')

-- Net Profit by Store (used in visual filter context)
-- Achieved by placing Store_Name on visual axis with [Net Profit] as value
-- Power BI automatically filters the measure to each store's context
```

### Power BI Dashboard
The dashboard provides an interactive 5-page view of sales performance, customer intelligence, channel profitability, and discount impact.

- Dashboard Features

> - KPIs (Gauge & Card Visuals)
> - Total Sales
> - Gross Profit
> - Net Profit Margin %
> - Average Transaction Value
> - Gross Potential Revenue
> - Sales Leakage
> - Total Orders
> - Total Revenue
> - Returned Revenue
> - Cancelled Revenue
> - Gross Profit
> - Online Revenue — Gross Potential
> - Actual Online Revenue
> - Total Customers
> - Total Active Customers
> - Total Inactive Customers
> - Average CLV
> - Repeat Purchase Rate %
> - Total Stores
> - Avg Discount %
> - Discount Revenue Impact %
> - Profit Before Discount
> - Net Profit
> - Total Discount Value

- Filters / Slicers
> - Order Status
> - Year
> - Sales Channel
> - Region
> - Customer Segment
> - Category

- Visuals — Page 1: Sales Overview

  
<img width="916" height="503" alt="Visuals 1" src="https://github.com/user-attachments/assets/a1314b95-282d-4c82-abdf-975ae84f9763" />

1. Total Revenue by Month and Year (Line Chart)
2. Revenue by Product Category (Donut Chart)
3. Revenue vs Gross Potential Revenue by Sales Channel (Clustered Bar Chart)
4. Total Revenue by Region (Bar Chart)
5. Total Sales (Gauge)
6. Gross Profit (Gauge)
7. Net Profit Margin % (Gauge)
8. Average Transaction Value (Gauge)
9. Gross Potential Revenue (Gauge)
10. Sales Leakage (Gauge)
11. Total Orders (Card)

- Visuals — Page 2: Sales Performance


<img width="917" height="501" alt="Visuals 2" src="https://github.com/user-attachments/assets/a9f038c1-30a2-4379-9422-9bb4a2ad7634" />


1. Revenue Lost by Returns and Cancellations by Category (Clustered Column Chart)
2. Net Profit by Sales Channel and Year (Clustered Column Chart)
3. Net Profit by Brand (Clustered Column Chart)
4. Average Transaction Value by Sales Channel (Bar Chart)
5. Gross Profit vs Total Revenue by Sales Channel (Clustered Bar Chart)
6. Total Revenue (Card)
7. Gross Profit (Card)
8. Returned Revenue (Card)
9. Cancelled Revenue (Card)
10. Online Revenue — Gross Potential (Card)
11. Actual Online Revenue (Card)
12. Revenue Leakage (Card)

- Visuals — Page 3: Customer Intelligence


<img width="900" height="489" alt="Visuals 3" src="https://github.com/user-attachments/assets/f4f08dfa-f893-4f6c-b41f-633d76983ad0" />


1. Net Profit by Loyalty Status (Pie Chart)
2. Total Revenue by Customer Segment (Column Chart)
3. Total Revenue by City (Clustered Bar Chart)
4. Total Revenue by Age Band (Clustered Column Chart)
5. Total Sales by Gender (Pie Chart)
6. Total Customers (Card)
7. Total Active Customers (Card)
8. Total Inactive Customers (Card)
9. Avg CLV (Card)
10. Repeat Purchase Rate % (Card)
11. Total Stores (Card)
12. Avg Discount % (Card)

- Visuals — Page 4: Channel Analysis

<img width="919" height="501" alt="Visuals 4" src="https://github.com/user-attachments/assets/6eebfc91-1c0b-4443-9d82-a7144c6f7185" />


1. Net Profit by Store Name (Clustered Bar Chart)
2. Total Revenue by Region (Clustered Column Chart)
3. Net Profit by Sales Channel (Donut Chart)
4. Total Customers (Card)
5. Total Active Customers (Card)
6. Total Inactive Customers (Card)
7. Avg CLV (Card)
8. Repeat Purchase Rate % (Card)
9. Total Stores (Card)
10. Avg Discount % (Card)

- Visuals — Page 5: Discount & Pricing Analysis


<img width="895" height="492" alt="Visuals 5" src="https://github.com/user-attachments/assets/3afd6a4f-bfb8-455f-8156-ef86d413f2e5" />


1. Total Discount Value by Category (Clustered Bar Chart)
2. Gross Profit vs Total Discount Value by Category (Bar Chart)
3. Total Revenue by Sales Channel (Clustered Column Chart)
4. Total Revenue by Discount Band (Clustered Column Chart)
5. Discount Revenue Impact % (Card)
6. Avg Discount % (Card)
7. Profit Before Discount (Card)
8. Net Profit (Card)
9. Total Discount Value (Card)

### Key Metrics
- Total Revenue: Tracked across all channels, categories and regions
- Total Orders: Full transaction volume across completed, returned and cancelled orders
- Gross Profit: Revenue after cost of goods sold
- Net Profit: Revenue after costs and discounts applied
- Net Profit Margin %: Profitability ratio across channels and categories
- Gross Potential Revenue: Maximum achievable revenue before returns and cancellations
- Sales Leakage: Revenue gap between gross potential and actual revenue
- Returned Revenue: Total value of returned orders
- Cancelled Revenue: Total value of cancelled orders
- Revenue Leakage: Combined impact of returns and cancellations on revenue
- Average Transaction Value: Average order value across sales channels
- Total Customers: Full customer base including active and inactive
- Total Active Customers: Customers with recent completed transactions
- Total Inactive Customers: Customers with no recent activity
- Average CLV: Average customer lifetime value across segments
- Repeat Purchase Rate %: Proportion of customers making more than one purchase
- Total Stores: Number of retail locations
- Avg Discount %: Average discount applied across transactions
- Total Discount Value: Total monetary value of discounts given
- Discount Revenue Impact %: Percentage of revenue lost to discounting
- Profit Before Discount: Revenue before discount deductions

### Key Insights
1. Revenue leakage from returns and cancellations is a critical business issue
- A significant portion of gross potential revenue is lost through returned and cancelled orders. This gap between potential and actual revenue — measured as Sales Leakage — represents a key area for operational improvement.

2. Sales channel profitability varies significantly
- Net profit and average transaction value differ considerably across Online, In-Store and Wholesale channels. Online revenue shows a measurable gap between gross potential and actual revenue, suggesting conversion and fulfilment challenges.

3. Customer segment and loyalty drive revenue concentration
- Premium and high-loyalty customers contribute disproportionately to total revenue and net profit. A small proportion of active customers generate the majority of repeat purchases, highlighting the importance of retention over acquisition.

4. Discounting erodes profitability across key categories
- High discount bands reduce net profit margins significantly. Categories receiving the highest total discount value do not always show proportionally higher revenue — suggesting discounting is not always driving incremental sales.

5. Regional performance is uneven across stores
- Revenue and net profit vary significantly by store and region. Certain stores consistently underperform on net profit despite reasonable revenue — indicating cost or returns issues at the store level.

6. Inactive customers represent a significant untapped opportunity
- A substantial proportion of the customer base is classified as inactive. Re-engagement of this segment through targeted campaigns could recover revenue without the cost of new customer acquisition.

7. Brand-level profitability reveals portfolio concentration risk
- Net profit is concentrated in a small number of top-performing brands. Heavy reliance on a few brands creates vulnerability to supply chain or demand disruption.

### Recommendations
1. Reduce revenue leakage from returns and cancellations
- Investigate the root causes of high return and cancellation rates by category and channel. Implement clearer product descriptions, size guides and cancellation windows to reduce avoidable losses.

2. Optimise online channel performance
- Close the gap between Online Gross Potential Revenue and Actual Online Revenue through improved checkout experience, targeted retargeting campaigns and abandoned cart recovery strategies.

3. Restructure discount strategy
- Replace blanket discount bands with targeted promotions based on customer segment and purchase history. Reserve high discounts for re-engagement campaigns rather than applying them broadly across all categories.

4. Invest in customer retention over acquisition
- With repeat purchase rates and CLV varying significantly by loyalty tier, resources should be directed towards retaining Gold and Silver loyalty customers rather than solely acquiring new ones.

5. Re-engage inactive customers
- Design a re-engagement campaign targeting inactive customers with personalised offers based on their previous purchase history. Even a modest reactivation rate would deliver meaningful revenue recovery.

6. Address underperforming stores
- Stores with high revenue but low net profit should be audited for returns rates, staffing costs and discount practices. Store-level profitability dashboards should be reviewed monthly by regional managers.

7. Diversify brand portfolio to reduce concentration risk
- Expand the product mix across mid-tier brands to reduce dependency on top performers. Negotiate better margin structures with high-volume brands to improve net profit across the portfolio.

### Dataset Source
The dataset was built for the RetailMax Group Ltd Enterprise Revenue Intelligence project, combining sales transaction data, customer profiles, product catalogues and store information across multiple sales channels and regions.
