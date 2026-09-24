# Global E-Commerce Sales & Profitability Optimization Analytics

IBM Data Analytics Internship Project

## Overview

This project analyzes the **Brazilian E-Commerce Public Dataset by Olist** using a reproducible Python pipeline. It covers KPI analysis, trends, category performance, regional freight burden, customer repeat behavior, delivery performance, statistical association, risk/opportunity evidence, and business recommendations.

## Dataset

Source: https://www.kaggle.com/datasets/olistbr/brazilian-ecommerce

The project expects the original CSV files in `data/raw/`.

Expected files:

- `olist_customers_dataset.csv`
- `olist_geolocation_dataset.csv`
- `olist_order_items_dataset.csv`
- `olist_order_payments_dataset.csv`
- `olist_order_reviews_dataset.csv`
- `olist_orders_dataset.csv`
- `olist_products_dataset.csv`
- `olist_sellers_dataset.csv`
- `product_category_name_translation.csv`

Do not upload the full raw dataset to GitHub unless its terms and repository size make that appropriate. The README provides the public source instead.

## Profitability Limitation

The Olist data does not provide complete COGS / operating cost information. The project therefore uses:

`Freight-Adjusted Contribution Proxy = Item Revenue - Freight Value`

This is **not** true accounting profit.

## Project Structure

```text
project/
├── project.py
├── requirements.txt
├── README.md
├── final_report.docx
├── .gitignore
├── data/
│   └── raw/
│       └── [Olist CSV files]
└── outputs/
    ├── analytical_master_table.csv
    ├── executive_summary.md
    ├── figures/
    └── tables/
```

## Setup on Windows

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Place the extracted Olist CSV files in `data/raw/`.

Run:

```bash
python project.py
```

The script generates the analytical tables, charts, executive summary and `final_report.docx`.

## Main Outputs

- `outputs/tables/kpis.csv` — executive KPIs
- `outputs/tables/data_dictionary.csv` — analytical data dictionary
- `outputs/tables/category_analysis.csv` — category-level analysis
- `outputs/tables/regional_analysis.csv` — state-level analysis
- `outputs/tables/customer_analysis.csv` — customer-level analysis
- `outputs/tables/delivery_analysis.csv` — delay group analysis
- `outputs/tables/driver_test_summary.csv` — statistical tests
- `outputs/tables/risk_analysis.csv` — evidence-based risk signals
- `outputs/tables/opportunity_analysis.csv` — evidence-based opportunities
- `outputs/figures/` — report-ready visualizations
- `final_report.docx` — completed internship report

## KPI Definitions

- **Total item revenue:** sum of item selling prices.
- **Total freight:** sum of freight charges.
- **AOV:** mean total order value among orders with item records.
- **Freight-to-revenue ratio:** freight / item revenue.
- **Contribution proxy:** item revenue - freight.
- **Late-delivery rate:** delivered orders whose actual delivery date is later than the estimated delivery date.
- **Repeat customer rate:** share of unique customers with more than one order.

## Interpretation Rules

Do not turn associations into causal claims. Do not call the contribution proxy accounting profit. Do not copy hypothetical findings from earlier planning documents; use the generated CSV outputs and charts.

## Internship Submission Checklist

- [ ] `project.py` runs from start to finish.
- [ ] `requirements.txt` installs successfully.
- [ ] `README.md` explains setup.
- [ ] `final_report.docx` contains actual numbers and charts.
- [ ] GitHub contains the required files.
- [ ] Raw dataset is referenced rather than unnecessarily committed.
- [ ] All findings match the generated tables.

