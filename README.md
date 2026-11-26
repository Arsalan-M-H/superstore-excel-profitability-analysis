# Superstore Retail Profitability Analysis | Excel

## Project Overview

This project investigates why a retail business can generate healthy sales while certain regions, product categories and subcategories record weak or negative profitability.

Using the Sample Superstore dataset, I conducted an end-to-end analysis in Microsoft Excel to investigate regional performance, discount strategies, product-level profitability and order-level returns. The objective was to identify the main areas of concern, understand potential contributing factors and provide evidence-supported recommendations that management could use to improve business performance.

The project includes the Excel analysis workbook and a management report containing the findings, supporting visualisations, limitations and prioritised recommendations.

## Business Problem

Although all four regions recorded positive overall profits, profitability varied substantially between regions and certain product groups consistently generated losses.

The analysis aimed to answer the following questions:

- Why are certain regions less profitable despite generating substantial sales?
- How do different discount levels relate to recorded profitability?
- Which product categories and subcategories contribute most to losses?
- Are there operational risks associated with order returns?
- What further information is required before management makes changes to pricing, inventory or operational processes?

## Tools and Skills Demonstrated

**Tools:** Microsoft Excel, PivotTables, Excel formulas, conditional formatting and charts.

**Analytical skills:**
- Data quality checks and validation
- Data cleaning and transformation
- Order-level and order-line-level analysis
- KPI development and interpretation
- Regional and product profitability analysis
- Discount-level analysis
- Data visualisation and business reporting
- Analytical reasoning and evidence-based recommendations
- Identification of data limitations and unsupported assumptions

**Excel techniques:** `XLOOKUP`, `IF`, `COUNTIF`, `COUNTIFS`, `SUMIFS`, `FILTER`, `UNIQUE`, PivotTables and calculated fields.

## Dataset and Methodology

The analysis uses the Sample Superstore dataset covering **January 2014 to December 2017**.

The workbook contains three source datasets:

| Dataset | Purpose |
|---|---|
| Orders | Transaction-level sales, profit, quantities, discounts, products and regional information |
| Returns | Identifies Order IDs flagged as returned |
| People | Regional management reference information |

The Orders dataset contains **9,993 order lines**. Since a single order can contain multiple products, order-line counts and distinct-order counts were treated separately.

The analysis followed these stages:

1. **Data validation:** Checked missing values, duplicate records, date consistency and transaction-level financial values.
2. **Data preparation:** Used Excel formulas and lookup functions to create supporting fields and identify orders appearing in the Returns dataset.
3. **Regional analysis:** Compared sales, recorded profit, profit margin, distinct orders, average order value and return rates.
4. **Product analysis:** Investigated profitability by category and subcategory across regions.
5. **Discount analysis:** Compared recorded financial performance across discount levels, focusing on loss-making product groups.
6. **Reporting:** Developed charts, supporting tables and a management report with prioritised recommendations.

Key metrics were calculated as follows:

- **Profit Margin:** Total recorded profit / Total recorded sales
- **Average Order Value:** Total recorded sales / Distinct orders
- **Order-Level Return Rate:** Distinct orders flagged as returned / Total distinct orders

## Key Findings

### 1. Regional Profitability

All four regions recorded positive overall profit margins, but their performance varied.

| Region | Profit Margin | Distinct Orders | Order-Level Return Rate |
|---|---:|---:|---:|
| Central | 7.92% | 1,175 | 3.32% |
| South | 11.93% | 822 | 2.92% |
| East | 13.49% | 1,401 | 3.14% |
| West | 14.94% | 1,611 | 11.73% |

Central recorded the lowest overall profit margin, while West had the highest recorded profit margin but also a substantially higher order-level return rate than the other regions.

### 2. High Discounts Were Associated With Significant Losses

The strongest recurring finding was the relationship between high discount levels and negative recorded profitability.

- Full-price transactions generally recorded strong profit margins.
- Lower discounts, including 10%, 15% and 20%, often remained profitable.
- No region recorded positive aggregate profit on transactions discounted by 30% or more.
- Transactions discounted above 20% generated approximately **$135,364 in net recorded losses** across the four regions.

Recorded profit from transactions discounted by 20% or less was approximately **$421,773**, compared with total recorded profit of approximately **$286,409** when all transactions were included.

These figures describe historical transaction groups. They do not demonstrate that removing high discounts would increase profit by the same amount, since customer demand and sales volumes could change.

### 3. Central's 80% Discounts Were a Major Contributor to Weak Profitability

Central recorded substantial losses from Appliances and Binders sold at an 80% discount.

Across these two subcategories, 300 heavily discounted order lines generated approximately **$30,539 in recorded losses**.

Excluding those transactions from the historical data increases Central's calculated profit margin from 7.92% to approximately 14.51%.

However, further information about promotional objectives, inventory clearance and expected financial outcomes would be required to determine whether these discounts were commercially justified.

### 4. Loss-Making Subcategories Appeared Across Multiple Regions

The analysis identified recurring profitability concerns in subcategories including:

- **Machines:** High discounts reduced profitability in Central, South and West.
- **Bookcases:** Heavy discounting was associated with negative aggregate profits in several regions.
- **Tables:** Recorded losses appeared in Central, East and South.
- **Appliances and Binders:** Central experienced substantial losses associated with 80% discounts.

Cross-regional comparisons highlighted that the same subcategory could be profitable at lower discounts but unprofitable at higher ones.

This supported further investigation into pricing and promotional strategies rather than treating every loss-making subcategory as an isolated issue.

### 5. West Recorded a Higher Order-Level Return Rate

West recorded an order-level return rate of **11.73%**, compared with approximately 3% in the other regions.

This represents a potential operational concern, but the Returns dataset does not identify the specific products returned, the reasons for returns or their financial impact.

Consequently, the analysis could not determine which products or processes were responsible for the elevated rate.

## Management Recommendations

### Priority 1 — Review Discount and Promotional Strategies

Investigate the objectives behind higher discounts, particularly whether they were intended to increase demand, support seasonal campaigns or clear discontinued inventory.

Introduce or review promotional forecasts covering expected sales volumes, revenue, profit margins and inventory reduction. Compare actual outcomes against the original objectives before adjusting discount policies.

### Priority 2 — Investigate West's Elevated Return Rate

Improve the collection of item-level return information, including product identifiers, return reasons, quantities, refund amounts and processing costs.

Compare return rates across historical periods to establish whether West's elevated rate is persistent and identify potential operational causes.

### Priority 3 — Assess Inventory Availability and Low-Volume Products

Investigate subcategories recording strong profit margins but relatively low unit sales.

Compare sales volumes against historical demand, inventory availability and stockout records to determine whether potential sales are being limited by supply shortages or simply reflect normal purchasing patterns.

Where physical stores exist, selected merchandising changes could also be tested and evaluated using sales and profit outcomes.

## Limitations

Several limitations affect how the findings should be interpreted.

**Order-level returns:** The Returns dataset flags entire Order IDs and does not identify which individual items were returned. As a result, product-level return rates and actual return-related financial losses cannot be reliably calculated.

**Recorded financial values:** It is not established whether the Sales and Profit fields already account for refunds or other return-related costs.

**Discount causality:** The analysis identifies associations between higher discount levels and lower profitability, but cannot establish how sales would have performed without those discounts.

**Missing operational information:** Inventory availability, promotional objectives, customer feedback and complete return-cost information were unavailable.

The recommendations therefore distinguish between observed findings and hypotheses requiring additional investigation.

## Repository Files

| File | Description |
|---|---|
| [Superstore_Analysis.xlsx](Superstore_Analysis.xlsx) | Excel workbook containing source data, supporting calculations, PivotTables and analysis |
| [Superstore_Management_Report.docx](Superstore_Management_Report.docx) | Management report containing detailed findings, visualisations, limitations and recommendations |

**How to explore the project:** Download the Excel workbook and open the `START HERE` worksheet for an overview of the project and navigation instructions. Read the management report for the full investigation and conclusions.

## Key Takeaways

This project strengthened my ability to approach business problems through structured data analysis, rather than focusing only on producing charts or summary statistics.

A particularly important lesson was understanding the difference between a technically correct calculation and a business conclusion that the data can support.

For example, an order flagged as returned does not necessarily identify which product was physically returned, and recorded losses from discounted transactions do not establish how much profit the business could recover by eliminating those discounts.

The project reinforced the importance of validating data, distinguishing transaction and order-level metrics, investigating financial performance in context, and communicating findings with appropriate limitations.

---

**Project type:** Independent Data Analyst portfolio project

**Tools:** Microsoft Excel

**Dataset:** Sample Superstore (2014–2017)

**Note:** This project uses a sample dataset for analytical and educational purposes. Findings and management recommendations are illustrative rather than conclusions about a real retail business.
