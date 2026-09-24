# Executive Analytics Summary

## Dataset
- Olist orders table: 99,441 orders.
- Item table: 112,650 order-item records.
- Purchase period: 2016-09-04 to 2018-10-17.

## Financial metric definition
The dataset does not provide complete product/operating cost information. The project therefore uses a **Freight-Adjusted Contribution Proxy = item revenue - freight value**. This is not accounting profit.

## Actual KPIs
- Total registered orders: 99,441
- Delivered orders: 96,478 (97.02%)
- Item revenue: BRL 13,591,643.70
- Freight: BRL 2,251,909.54
- Contribution proxy: BRL 11,339,734.16
- Freight-to-revenue ratio: 16.57%
- Average order value incl. freight: BRL 160.58
- Average delivery time: 12.56 days
- Late delivery rate: 8.11%
- Average review score: 4.09
- Repeat customer rate: 3.12%

## Evidence-based findings
1. **Revenue concentration:** health_beauty is the highest-revenue product category at BRL 1,258,681.34.
2. **Freight-adjusted contribution:** watches_gifts has the largest contribution proxy at BRL 1,104,469.75, with a freight ratio of 8.34%.
3. **Regional scale:** SP generates the largest state-level revenue at BRL 5,202,955.05 across 41,375 orders.
4. **Regional freight burden:** Among states with at least 500 orders, MA has the highest freight-to-revenue ratio at 26.35%.
5. **Customer experience:** 7,826 delivered orders were late in the delay analysis; delivery delay and review score show a Spearman association of -0.176 (p < 0.001).

## Customer retention finding
Only 3.12% of unique customers are repeat customers in the historical data, while repeat-customer orders represent 6.38% of all orders. This supports investigating retention and second-purchase conversion rather than assuming high customer lifetime value.

## Interpretation and limitations
- The financial proxy excludes COGS and other operating costs, so it must not be interpreted as true profit.
- High freight ratio is evidence of freight burden, not proof that freight alone causes weak profitability.
- The delay/review relationship is statistical association, not proof of causation.
- Very small states/categories should be interpreted cautiously; the project uses sample-size thresholds for operational risk flags.
- The dataset is historical (2016-2018) and may not represent current e-commerce operations.