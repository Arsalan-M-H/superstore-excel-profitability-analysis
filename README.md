# Superstore Retail Profitability and Scenario Analysis | Excel

## Project Overview

This project investigates why a retail business can generate healthy sales while certain regions, product categories and subcategories record weak or negative profitability.

Using the Sample Superstore dataset, I conducted an end-to-end analysis in Microsoft Excel to investigate regional performance, discount strategies, product profitability and order-level returns. The project was further extended to include financial scenario modelling and price elasticity sensitivity analysis to explore how alternative discount strategies could affect demand, revenue and profitability.

The objective was to identify areas of concern, investigate potential contributing factors, evaluate alternative business decisions and provide evidence-supported recommendations for management.

## Business Questions

- Why do certain regions record weaker profitability despite substantial sales?
- Which categories and subcategories contribute most to recorded losses?
- How are different discount levels associated with profitability?
- Which operational concerns require further investigation?
- How could reducing high discounts affect revenue, sales volume and profit?
- How sensitive are simulated financial outcomes to different assumptions about customer price sensitivity?

## Tools and Skills Demonstrated

### Technical Skills

- Microsoft Excel
- Data cleaning, transformation and validation
- PivotTables and calculated fields
- Dynamic array formulas
- XLOOKUP and conditional aggregation
- FILTER, UNIQUE, CHOOSECOLS and XMATCH
- COUNTIF, COUNTIFS and SUMIFS
- Conditional formatting and data visualisation
- Interactive scenario modelling and financial calculations

### Analytical and Business Skills

- KPI development and interpretation
- Regional and product profitability analysis
- Order-level and order-line-level analysis
- Commercial and promotional performance analysis
- Financial scenario modelling
- Price elasticity of demand
- Sensitivity analysis
- Extreme-value testing
- Identification of confounding variables
- Evaluation of modelling assumptions and data limitations
- Evidence-based business recommendations
- Management reporting

## Dataset

The project uses the Sample Superstore dataset covering January 2014 to December 2017.

| Dataset | Description |
|---|---|
| Orders | Order-line sales, profit, quantities, discounts, products and regional information |
| Returns | Order IDs flagged as returned |
| People | Regional management reference information |

The dataset contains 9,993 order lines representing 5,009 distinct orders. Since a single order can contain multiple products, order-line counts and distinct-order counts were treated separately.

## Analytical Methodology

The project was completed through the following stages:

1. **Business problem definition:** Established the commercial questions and analytical objectives.
2. **Data validation:** Investigated missing values, duplicate records, date consistency and transaction-level financial values.
3. **Data preparation:** Created supporting fields and identified orders appearing in the Returns dataset.
4. **Regional analysis:** Compared recorded sales, profit, profit margins, distinct orders, average order values and return rates.
5. **Product analysis:** Investigated category and subcategory profitability across regions.
6. **Discount analysis:** Examined financial performance at different discount levels and identified recurring loss-making product groups.
7. **Scenario modelling:** Simulated the financial impact of halving discounts greater than 20%.
8. **Sensitivity analysis:** Evaluated simulated profitability under different assumptions about price elasticity.
9. **Reporting:** Developed supporting visualisations, documented analytical limitations and prepared a management report with prioritised recommendations.

### KPI Definitions

- **Profit Margin:** Total recorded profit / Total recorded sales
- **Average Order Value:** Total recorded sales / Distinct orders
- **Order-Level Return Rate:** Distinct orders flagged as returned / Total distinct orders

## Key Findings

### 1. Regional Performance

| Region | Distinct Orders | Profit Margin | Order-Level Return Rate |
|---|---:|---:|---:|
| Central | 1,175 | 7.92% | 3.32% |
| East | 1,401 | 13.49% | 3.14% |
| South | 822 | 11.93% | 2.92% |
| West | 1,611 | 14.94% | 11.73% |

All four regions recorded positive overall profits. However, Central recorded the weakest profit margin, while West achieved the highest profit margin alongside a considerably higher order-level return rate.

### 2. Discount Profitability

Higher discounts were consistently associated with weaker recorded profitability.

- Full-price transactions generally recorded strong profit margins.
- Lower discount levels often remained profitable.
- No region recorded positive aggregate profit on transactions discounted by 30% or more.
- Transactions discounted above 20% generated approximately $135,364 in net recorded losses.

| Historical Transaction Group | Recorded Profit |
|---|---:|
| Discounts of 20% or less | $421,773.08 |
| Discounts greater than 20% | -$135,364.00 |
| Overall business | $286,409.08 |

These figures describe historical financial performance and do not establish that eliminating high discounts would recover the recorded losses.

### 3. Central's High-Discount Transactions

Central recorded substantial losses from Appliances and Binders sold at an 80% discount.

Across 300 order lines, these transactions generated approximately $30,539 in recorded losses.

Excluding these transactions from the historical data increases Central's calculated profit margin from 7.92% to approximately 14.51%.

Further information about promotional objectives, stock clearance and expected financial outcomes would be required before determining whether these discounts were commercially justified.

### 4. Cross-Regional Product Performance

Recurring profitability concerns were identified in several subcategories:

- Machines
- Bookcases
- Tables
- Appliances
- Binders

Cross-regional comparisons demonstrated that the same subcategory could record positive profitability at lower discount levels and substantial losses at higher discount levels.

### 5. West's Elevated Return Rate

West recorded an order-level return rate of 11.73%, compared with approximately 3% across the other regions.

This represents a potential operational concern. However, the Returns dataset does not identify the individual products returned, reasons for returns, refund amounts or associated operational costs.

Further investigation and improved data collection would therefore be required before implementing corrective measures.

## Scenario Analysis — Reducing High Discounts

Following the historical discount analysis, I developed an interactive scenario model to investigate the potential financial impact of halving all discounts greater than 20%.

For example, a historical discount of 80% becomes 40%, while a historical discount of 50% becomes 25%.

The scenario model covers 1,392 historical order lines. Transactions discounted by 20% or less remain unchanged.

### Scenario Methodology

The original undiscounted unit price was reconstructed from historical sales, quantities and discount levels.

Original Unit Price = Historical Sales / (Historical Quantity × (1 − Historical Discount))

The scenario selling price was then calculated by applying half the original discount.

Scenario Selling Price = Original Unit Price × (1 − Scenario Discount)

To avoid assuming that customers would purchase the same quantity at higher prices, a constant-elasticity demand model was introduced.

Scenario Quantity = Historical Quantity × (Scenario Selling Price / Historical Selling Price) ^ Assumed Elasticity

Historical unit costs were inferred from recorded sales, profit and quantities.

Estimated Unit Cost = (Historical Sales − Historical Profit) / Historical Quantity

Scenario revenue and profit were subsequently calculated using the simulated quantities and estimated unit costs.

### Price Elasticity Sensitivity Analysis

Four hypothetical price elasticity assumptions were tested, ranging from relatively inelastic to relatively elastic demand.

| Assumed Elasticity | Simulated Profit on Affected Transactions | Simulated Total Business Profit |
|---|---:|---:|
| Historical baseline | -$135,364.00 | $286,409.08 |
| -0.5 | $38,474.99 | $460,248.07 |
| -1.0 | $32,022.06 | $453,795.14 |
| -1.5 | $26,972.07 | $448,745.15 |
| -2.0 | $22,937.60 | $444,710.67 |

All four scenarios produced higher simulated profitability than the historical baseline, with improvements ranging from approximately $158,302 to $173,839.

As the assumed price sensitivity increased, simulated profitability decreased because customers were modelled as purchasing fewer units at the higher selling prices.

### Scenario Validation and Limitations

The scenario workbook includes an editable elasticity input, allowing users to investigate how different demand assumptions affect financial outcomes.

Extreme-value testing was also conducted to understand the behaviour of the model when demand approaches zero. This exposed an important structural limitation: because transactions outside the affected discount groups remain unchanged and costs associated with unsold units are excluded from simulated profit, eliminating historically loss-making sales can still produce an apparent financial improvement.

However, this does not account for inventory already purchased, storage costs, stock clearance objectives, product obsolescence or the potential impact of promotional activity on other purchases.

The elasticities used in this project are hypothetical and were not estimated from observed customer behaviour.

The results should therefore be interpreted as illustrative financial scenarios rather than forecasts or causal estimates of the impact of changing prices.

## Management Recommendations

### Priority 1 — Review Discount and Promotional Strategies

Investigate the objectives behind higher discounts, including seasonal campaigns, customer acquisition and inventory clearance.

The scenario results support further investigation of alternative discount strategies. However, management should validate customer price sensitivity, promotional objectives, inventory implications and wider purchasing behaviour before implementing changes.

Promotional performance should be evaluated against forecasts for revenue, units sold, profit, profit margin and relevant inventory objectives.

### Priority 2 — Investigate West's Elevated Return Rate

Improve the collection of item-level return information, including product identifiers, return reasons, quantities, refund amounts and processing costs.

Compare return rates across historical periods to determine whether West's elevated rate represents a persistent operational issue.

### Priority 3 — Assess Inventory Availability and Low-Volume Products

Investigate subcategories with strong profit margins but relatively low unit sales.

Compare historical demand, inventory availability and stockout records to determine whether low volumes reflect expected purchasing behaviour or potential supply limitations.

Where applicable, controlled merchandising experiments could also be conducted to evaluate opportunities for additional sales and profitability.

## Analytical Limitations

- Returns are recorded at order level and cannot be reliably attributed to individual products.
- Recorded sales and profit may not reflect the complete financial impact of refunds.
- Promotional objectives, marketing expenditure, stock availability and customer feedback are unavailable.
- Historical price and quantity relationships may be influenced by seasonality, customer traffic, product exposure, conversion rates and other confounding variables.
- The scenario model assumes a constant price elasticity shared across affected transactions.
- Estimated unit costs are held constant and modelled costs vary proportionally with quantities sold.
- Purchases outside the affected discount groups are assumed to remain unchanged.
- The scenario analysis does not establish actual customer price elasticity or predict future business performance.

## Repository Files

| File | Description |
|---|---|
| Excel analysis workbook (.xlsx) | Source data, PivotTables, supporting analysis, visualisations and interactive scenario modelling |
| Management report (.docx) | Detailed findings, scenario results, limitations and prioritised management recommendations |
| README.md | Project overview, methodology and key findings |

To explore the project, download the Excel workbook and begin with the START HERE worksheet. The Scenario Analysis worksheet contains the interactive elasticity model, sensitivity comparison and underlying transaction-level calculations.

## Learning Outcomes

This project strengthened my ability to approach business problems through structured analysis and to distinguish technically correct calculations from commercially defensible conclusions.

In addition to developing my Excel proficiency, I explored price elasticity, financial scenario modelling, sensitivity analysis and the importance of challenging modelling assumptions.

A particularly important learning outcome was understanding that a model can produce mathematically correct results without establishing what would actually happen in the real world.

The project reinforced the importance of validating calculations, considering alternative explanations, communicating uncertainty and recommending actions that are supported by the available evidence.

---

**Project Type:** Independent Data Analyst portfolio project

**Tools:** Microsoft Excel

**Dataset:** Sample Superstore (2014–2017)

**Note:** This project uses a sample dataset for analytical and educational purposes. Its findings and recommendations are illustrative, and the scenario results should not be interpreted as forecasts.
