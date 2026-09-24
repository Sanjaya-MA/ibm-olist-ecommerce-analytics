"""
IBM Data Analytics Internship Project
Global E-Commerce Sales & Profitability Optimization Analytics

Source dataset: Brazilian E-Commerce Public Dataset by Olist
https://www.kaggle.com/datasets/olistbr/brazilian-ecommerce

This project intentionally uses the real Olist CSV files. It does not generate
synthetic data or hard-code business findings.

Financial limitation:
Olist provides selling price and freight value, but not full product/operating
costs. The project therefore reports a Freight-Adjusted Contribution Proxy:
    contribution_proxy = item_revenue - freight_value
This is not accounting profit or true net profit.
"""

from __future__ import annotations

import math
import sys
from pathlib import Path
from typing import Dict, Iterable, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy.stats import mannwhitneyu, spearmanr
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


PROJECT_ROOT = Path(__file__).resolve().parent
RAW_DIR = PROJECT_ROOT / "data" / "raw"
OUTPUT_DIR = PROJECT_ROOT / "outputs"
FIG_DIR = OUTPUT_DIR / "figures"
TABLE_DIR = OUTPUT_DIR / "tables"
REPORT_PATH = PROJECT_ROOT / "final_report.docx"

DATASET_URL = "https://www.kaggle.com/datasets/olistbr/brazilian-ecommerce"


# -----------------------------------------------------------------------------
# Utilities
# -----------------------------------------------------------------------------

def ensure_dirs() -> None:
    for path in (RAW_DIR, OUTPUT_DIR, FIG_DIR, TABLE_DIR):
        path.mkdir(parents=True, exist_ok=True)


def require_file(filename: str) -> Path:
    path = RAW_DIR / filename
    if not path.exists():
        raise FileNotFoundError(
            f"Required dataset file not found: {path}\n"
            f"Extract the Olist dataset into {RAW_DIR} before running the project."
        )
    return path


def fmt_currency(value: float) -> str:
    if pd.isna(value):
        return "N/A"
    return f"BRL {value:,.2f}"


def fmt_pct(value: float, digits: int = 2) -> str:
    if pd.isna(value):
        return "N/A"
    return f"{value * 100:.{digits}f}%"


def fmt_num(value: float, digits: int = 0) -> str:
    if pd.isna(value):
        return "N/A"
    return f"{value:,.{digits}f}"


def savefig(fig: plt.Figure, name: str) -> Path:
    path = FIG_DIR / name
    fig.savefig(path, dpi=170, bbox_inches="tight")
    plt.close(fig)
    return path


def set_repeat_table_header(row) -> None:
    trPr = row._tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    tblHeader.set(qn("w:val"), "true")
    trPr.append(tblHeader)


def shade_cell(cell, fill: str = "D9EAF7") -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def add_table(doc: Document, df: pd.DataFrame, title: str, max_rows: int = 12) -> None:
    if df.empty:
        doc.add_paragraph(f"{title}: no rows available.")
        return
    doc.add_heading(title, level=3)
    view = df.head(max_rows).copy()
    table = doc.add_table(rows=1, cols=len(view.columns))
    table.style = "Table Grid"
    table.autofit = True
    hdr = table.rows[0]
    set_repeat_table_header(hdr)
    for i, col in enumerate(view.columns):
        hdr.cells[i].text = str(col)
        shade_cell(hdr.cells[i])
        hdr.cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for par in hdr.cells[i].paragraphs:
            for run in par.runs:
                run.font.size = Pt(8.5)
                run.bold = True
    for _, row in view.iterrows():
        cells = table.add_row().cells
        for i, value in enumerate(row.tolist()):
            if isinstance(value, (float, np.floating)):
                text = f"{value:,.3f}"
            elif isinstance(value, (int, np.integer)):
                text = f"{value:,}"
            else:
                text = str(value)
            cells[i].text = text
            cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for par in cells[i].paragraphs:
                par.paragraph_format.space_after = Pt(0)
                for run in par.runs:
                    run.font.size = Pt(8.5)
    # Prevent individual rows from breaking across pages.
    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        cantSplit = OxmlElement("w:cantSplit")
        trPr.append(cantSplit)
    doc.add_paragraph("")


def add_figure(doc: Document, path: Path, caption: str) -> None:
    if path.exists():
        doc.add_picture(str(path), width=Inches(6.5))
        p = doc.paragraphs[-1]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap = doc.add_paragraph(caption)
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.runs[0].italic = True


# -----------------------------------------------------------------------------
# Data loading and cleaning
# -----------------------------------------------------------------------------

def load_data() -> Dict[str, pd.DataFrame]:
    files = {
        "customers": "olist_customers_dataset.csv",
        "geolocation": "olist_geolocation_dataset.csv",
        "items": "olist_order_items_dataset.csv",
        "payments": "olist_order_payments_dataset.csv",
        "reviews": "olist_order_reviews_dataset.csv",
        "orders": "olist_orders_dataset.csv",
        "products": "olist_products_dataset.csv",
        "sellers": "olist_sellers_dataset.csv",
        "translation": "product_category_name_translation.csv",
    }
    data: Dict[str, pd.DataFrame] = {}
    for key, filename in files.items():
        path = require_file(filename)
        data[key] = pd.read_csv(path)
    return data


def parse_dates(data: Dict[str, pd.DataFrame]) -> None:
    date_columns = {
        "orders": [
            "order_purchase_timestamp",
            "order_approved_at",
            "order_delivered_carrier_date",
            "order_delivered_customer_date",
            "order_estimated_delivery_date",
        ],
        "items": ["shipping_limit_date"],
        "reviews": ["review_creation_date", "review_answer_timestamp"],
    }
    for table_name, cols in date_columns.items():
        for col in cols:
            if col in data[table_name].columns:
                data[table_name][col] = pd.to_datetime(data[table_name][col], errors="coerce")


def inspect_raw_data(data: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    records = []
    key_candidates = {
        "customers": "customer_id",
        "geolocation": "geolocation_zip_code_prefix",
        "items": "order_id",
        "payments": "order_id",
        "reviews": "review_id",
        "orders": "order_id",
        "products": "product_id",
        "sellers": "seller_id",
        "translation": "product_category_name",
    }
    for name, df in data.items():
        key = key_candidates.get(name)
        records.append(
            {
                "table": name,
                "rows": len(df),
                "columns": len(df.columns),
                "exact_duplicate_rows": int(df.duplicated().sum()),
                "missing_cells": int(df.isna().sum().sum()),
                "primary_key_candidate": key or "",
                "unique_key_values": int(df[key].nunique(dropna=True)) if key in df.columns else np.nan,
                "column_names": ", ".join(df.columns),
            }
        )
    result = pd.DataFrame(records)
    result.to_csv(TABLE_DIR / "raw_table_inventory.csv", index=False)
    return result


def clean_data(data: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    data = {k: v.copy() for k, v in data.items()}
    parse_dates(data)

    # Remove exact duplicates only; legitimate multi-item order rows remain intact.
    for df in data.values():
        df.drop_duplicates(inplace=True)

    # Product text/category fields: preserve missingness as an explicit category.
    data["products"]["product_category_name"] = data["products"]["product_category_name"].fillna("unknown")

    # Numeric coercion for financial and review fields.
    for col in ["price", "freight_value"]:
        data["items"][col] = pd.to_numeric(data["items"][col], errors="coerce")
    data["payments"]["payment_value"] = pd.to_numeric(data["payments"]["payment_value"], errors="coerce")
    data["reviews"]["review_score"] = pd.to_numeric(data["reviews"]["review_score"], errors="coerce")

    # Core economic records need price and freight to calculate contribution proxy.
    data["items"] = data["items"].dropna(subset=["order_id", "price", "freight_value"])

    # Reviews can have more than one review record per order. Aggregate at order level.
    # Do not fabricate scores for orders without reviews.
    data["reviews"] = data["reviews"].dropna(subset=["order_id", "review_score"])

    return data


def create_data_dictionary(data: Dict[str, pd.DataFrame], analytical: pd.DataFrame) -> pd.DataFrame:
    source_map = {}
    for table_name, df in data.items():
        for col in df.columns:
            source_map.setdefault(col, table_name)

    descriptions = {
        "order_id": "Unique order identifier.",
        "customer_id": "Customer identifier for the order record.",
        "customer_unique_id": "Customer identifier stable across repeat orders.",
        "order_status": "Current status recorded for the order.",
        "order_purchase_timestamp": "Timestamp at which the order was placed.",
        "order_delivered_customer_date": "Timestamp when the order was delivered to the customer.",
        "order_estimated_delivery_date": "Estimated delivery date recorded for the order.",
        "item_revenue": "Sum of item selling prices for the order.",
        "freight_value": "Sum of item freight charges for the order.",
        "total_order_value": "Item revenue plus freight charges.",
        "contribution_proxy": "Item revenue minus freight; proxy only, not accounting profit.",
        "freight_ratio": "Freight divided by item revenue.",
        "item_count": "Number of item rows in the order.",
        "product_count": "Number of distinct products in the order.",
        "purchase_month": "Calendar month of purchase.",
        "purchase_year": "Calendar year of purchase.",
        "delivery_days": "Days from purchase to customer delivery.",
        "delivery_delay_days": "Actual delivery date minus estimated delivery date in days.",
        "delivered_late": "True when delivery occurred after the estimated delivery date.",
        "review_score": "Average available review score for the order.",
        "customer_type": "Repeat when the customer has more than one order; otherwise one-time.",
        "customer_order_count": "Number of orders associated with the unique customer identifier.",
        "dominant_payment_type": "Payment type with the greatest aggregated payment value for the order.",
        "payment_value": "Total recorded payment value for the order.",
        "state": "Customer state.",
    }
    records = []
    for col in analytical.columns:
        records.append(
            {
                "column_name": col,
                "data_type": str(analytical[col].dtype),
                "description": descriptions.get(col, "Derived analytical field or source attribute."),
                "source_table": source_map.get(col, "derived"),
                "missing_count": int(analytical[col].isna().sum()),
                "missing_percentage": round(float(analytical[col].isna().mean() * 100), 2),
                "unique_count": int(analytical[col].nunique(dropna=True)),
                "analytical_role": (
                    "Identifier" if col.endswith("_id") or col in {"order_id", "customer_id"}
                    else "Metric" if pd.api.types.is_numeric_dtype(analytical[col])
                    else "Dimension / Feature"
                ),
            }
        )
    result = pd.DataFrame(records)
    result.to_csv(TABLE_DIR / "data_dictionary.csv", index=False)
    return result


# -----------------------------------------------------------------------------
# Analytical model
# -----------------------------------------------------------------------------

def build_order_item_aggregate(items: pd.DataFrame) -> pd.DataFrame:
    result = (
        items.groupby("order_id", as_index=False)
        .agg(
            item_revenue=("price", "sum"),
            freight_value=("freight_value", "sum"),
            item_count=("order_item_id", "count"),
            product_count=("product_id", "nunique"),
        )
    )
    result["total_order_value"] = result["item_revenue"] + result["freight_value"]
    result["contribution_proxy"] = result["item_revenue"] - result["freight_value"]
    result["freight_ratio"] = np.where(
        result["item_revenue"] > 0,
        result["freight_value"] / result["item_revenue"],
        np.nan,
    )
    return result


def build_payment_aggregate(payments: pd.DataFrame) -> pd.DataFrame:
    value = payments.groupby("order_id", as_index=False).agg(payment_value=("payment_value", "sum"))
    by_type = (
        payments.groupby(["order_id", "payment_type"], as_index=False)["payment_value"]
        .sum()
        .sort_values(["order_id", "payment_value"], ascending=[True, False])
    )
    dominant = by_type.drop_duplicates("order_id", keep="first").rename(
        columns={"payment_type": "dominant_payment_type"}
    )[["order_id", "dominant_payment_type"]]
    return value.merge(dominant, on="order_id", how="left")


def build_review_aggregate(reviews: pd.DataFrame) -> pd.DataFrame:
    return (
        reviews.groupby("order_id", as_index=False)
        .agg(
            review_score=("review_score", "mean"),
            review_count=("review_id", "nunique"),
        )
    )


def build_analytical_master(data: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    orders = data["orders"].copy()
    customers = data["customers"][["customer_id", "customer_unique_id", "customer_city", "customer_state"]].copy()

    item_agg = build_order_item_aggregate(data["items"])
    payment_agg = build_payment_aggregate(data["payments"])
    review_agg = build_review_aggregate(data["reviews"])

    # Order-level master retains all Olist orders; financial fields are null for
    # orders that have no corresponding item rows.
    master = orders.merge(item_agg, on="order_id", how="left", validate="one_to_one")
    master = master.merge(customers, on="customer_id", how="left", validate="one_to_one")
    master = master.merge(payment_agg, on="order_id", how="left", validate="one_to_one")
    master = master.merge(review_agg, on="order_id", how="left", validate="one_to_one")

    master["purchase_month"] = master["order_purchase_timestamp"].dt.to_period("M").astype(str)
    master["purchase_year"] = master["order_purchase_timestamp"].dt.year
    master["delivery_days"] = (
        master["order_delivered_customer_date"] - master["order_purchase_timestamp"]
    ).dt.total_seconds() / 86400
    master["delivery_delay_days"] = (
        master["order_delivered_customer_date"] - master["order_estimated_delivery_date"]
    ).dt.total_seconds() / 86400
    master["delivered_late"] = master["delivery_delay_days"] > 0

    # Customer repeat classification is based on order history using stable customer ID.
    order_counts = master.groupby("customer_unique_id")["order_id"].nunique(dropna=True)
    master["customer_order_count"] = master["customer_unique_id"].map(order_counts)
    master["customer_type"] = np.where(
        master["customer_order_count"] > 1, "Repeat", "One-time"
    )

    master.to_csv(OUTPUT_DIR / "analytical_master_table.csv", index=False)
    return master


def build_item_category_table(data: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    items = data["items"].copy()
    products = data["products"][
        [
            "product_id",
            "product_category_name",
            "product_photos_qty",
            "product_weight_g",
            "product_length_cm",
            "product_height_cm",
            "product_width_cm",
        ]
    ].copy()
    translation = data["translation"].copy()

    result = items.merge(products, on="product_id", how="left", validate="many_to_one")
    result = result.merge(translation, on="product_category_name", how="left", validate="many_to_one")
    result["category"] = result["product_category_name_english"].fillna(result["product_category_name"])
    result["category"] = result["category"].fillna("unknown")
    result["contribution_proxy"] = result["price"] - result["freight_value"]
    return result


# -----------------------------------------------------------------------------
# Analytics
# -----------------------------------------------------------------------------

def calculate_kpis(master: pd.DataFrame) -> pd.DataFrame:
    fin = master[master["item_revenue"].notna()].copy()
    delivered = master[master["order_status"].eq("delivered")].copy()
    delivered_complete = delivered.dropna(subset=["order_delivered_customer_date", "order_estimated_delivery_date"])
    customers = master["customer_unique_id"].nunique(dropna=True)

    total_revenue = fin["item_revenue"].sum()
    total_freight = fin["freight_value"].sum()
    contribution = fin["contribution_proxy"].sum()

    rows = [
        ("Total registered orders", len(master), "count", "All rows in the Olist orders table."),
        ("Orders with item records", fin["order_id"].nunique(), "count", "Orders with item-level financial records."),
        ("Delivered orders", int((master["order_status"] == "delivered").sum()), "count", "Order status = delivered."),
        ("Unique customers", customers, "count", "Distinct customer_unique_id."),
        ("Total item revenue", total_revenue, "BRL", "Sum of item selling prices."),
        ("Total freight", total_freight, "BRL", "Sum of freight_value."),
        ("Freight-adjusted contribution proxy", contribution, "BRL", "Item revenue minus freight only; not accounting profit."),
        ("Contribution proxy margin", contribution / total_revenue if total_revenue else np.nan, "ratio", "Contribution proxy / item revenue."),
        ("Average order value incl. freight", fin["total_order_value"].mean(), "BRL", "Average total order value among orders with item records."),
        ("Freight-to-revenue ratio", total_freight / total_revenue if total_revenue else np.nan, "ratio", "Freight / item revenue."),
        ("Average delivery days", delivered_complete["delivery_days"].mean(), "days", "Delivered orders with valid purchase and delivery dates."),
        ("Late delivery rate", delivered_complete["delivered_late"].mean(), "ratio", "Delivered orders delivered after the estimate."),
        ("Average review score", master["review_score"].mean(), "score", "Mean of available order-level review scores."),
        ("Repeat customer rate", (master.groupby("customer_unique_id")["order_id"].nunique() > 1).mean(), "ratio", "Share of unique customers with more than one order."),
        ("Order share from repeat customers", (master["customer_type"] == "Repeat").mean(), "ratio", "Share of all orders linked to repeat customers."),
    ]
    result = pd.DataFrame(rows, columns=["kpi", "value", "unit", "definition"])
    result.to_csv(TABLE_DIR / "kpis.csv", index=False)
    return result


def analyze_category(item_category: pd.DataFrame) -> pd.DataFrame:
    result = (
        item_category.groupby("category", dropna=False)
        .agg(
            items=("order_item_id", "count"),
            orders=("order_id", "nunique"),
            revenue=("price", "sum"),
            freight=("freight_value", "sum"),
            avg_item_price=("price", "mean"),
            avg_freight=("freight_value", "mean"),
            avg_review_score=("review_score", "mean") if "review_score" in item_category.columns else ("price", "mean"),
        )
        .reset_index()
    )
    result["contribution_proxy"] = result["revenue"] - result["freight"]
    result["margin_proxy"] = np.where(result["revenue"] != 0, result["contribution_proxy"] / result["revenue"], np.nan)
    result["freight_ratio"] = np.where(result["revenue"] != 0, result["freight"] / result["revenue"], np.nan)
    result["revenue_per_order_containing_category"] = result["revenue"] / result["orders"]
    result = result.sort_values("revenue", ascending=False)
    result.to_csv(TABLE_DIR / "category_analysis.csv", index=False)
    return result


def analyze_region(master: pd.DataFrame) -> pd.DataFrame:
    fin = master[master["item_revenue"].notna()].copy()
    result = (
        fin.groupby("customer_state", dropna=False)
        .agg(
            orders=("order_id", "nunique"),
            revenue=("item_revenue", "sum"),
            freight=("freight_value", "sum"),
            contribution_proxy=("contribution_proxy", "sum"),
            avg_order_value=("total_order_value", "mean"),
            avg_delivery_days=("delivery_days", "mean"),
            late_rate=("delivered_late", "mean"),
            avg_review_score=("review_score", "mean"),
        )
        .reset_index()
    )
    result["freight_ratio"] = np.where(result["revenue"] != 0, result["freight"] / result["revenue"], np.nan)
    result["margin_proxy"] = np.where(result["revenue"] != 0, result["contribution_proxy"] / result["revenue"], np.nan)
    result = result.sort_values("revenue", ascending=False)
    result.to_csv(TABLE_DIR / "regional_analysis.csv", index=False)
    return result


def analyze_customer(master: pd.DataFrame) -> pd.DataFrame:
    fin = master[master["item_revenue"].notna()].copy()
    result = (
        fin.groupby("customer_unique_id", dropna=False)
        .agg(
            orders=("order_id", "nunique"),
            revenue=("item_revenue", "sum"),
            freight=("freight_value", "sum"),
            contribution_proxy=("contribution_proxy", "sum"),
            avg_order_value=("total_order_value", "mean"),
            avg_review_score=("review_score", "mean"),
            customer_state=("customer_state", "first"),
        )
        .reset_index()
    )
    result["customer_type"] = np.where(result["orders"] > 1, "Repeat", "One-time")
    result["contribution_margin_proxy"] = np.where(result["revenue"] != 0, result["contribution_proxy"] / result["revenue"], np.nan)
    result.to_csv(TABLE_DIR / "customer_analysis.csv", index=False)
    return result


def analyze_monthly(master: pd.DataFrame) -> pd.DataFrame:
    fin = master[master["item_revenue"].notna()].copy()
    result = (
        fin.groupby("purchase_month", dropna=False)
        .agg(
            orders=("order_id", "nunique"),
            revenue=("item_revenue", "sum"),
            freight=("freight_value", "sum"),
            contribution_proxy=("contribution_proxy", "sum"),
        )
        .reset_index()
    )
    result["contribution_margin_proxy"] = np.where(result["revenue"] != 0, result["contribution_proxy"] / result["revenue"], np.nan)
    result.to_csv(TABLE_DIR / "monthly_analysis.csv", index=False)
    return result


def analyze_delivery(master: pd.DataFrame) -> pd.DataFrame:
    d = master[
        master["order_status"].eq("delivered")
        & master["order_delivered_customer_date"].notna()
        & master["order_estimated_delivery_date"].notna()
    ].copy()
    d["delay_group"] = pd.cut(
        d["delivery_delay_days"],
        bins=[-np.inf, 0, 3, np.inf],
        labels=["On time / early", "1-3 days late", ">3 days late"],
        right=True,
    )
    result = (
        d.groupby("delay_group", observed=True)
        .agg(
            orders=("order_id", "nunique"),
            average_delivery_days=("delivery_days", "mean"),
            average_delay_days=("delivery_delay_days", "mean"),
            average_review_score=("review_score", "mean"),
        )
        .reset_index()
    )
    total_delivered = len(d)
    result["order_share"] = result["orders"] / total_delivered if total_delivered else np.nan
    result.to_csv(TABLE_DIR / "delivery_analysis.csv", index=False)
    return result


def analyze_drivers(master: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, float]]:
    d = master[
        master["order_status"].eq("delivered")
        & master["order_delivered_customer_date"].notna()
        & master["order_estimated_delivery_date"].notna()
        & master["review_score"].notna()
    ].copy()

    numeric_cols = [
        "item_revenue",
        "freight_value",
        "total_order_value",
        "contribution_proxy",
        "freight_ratio",
        "item_count",
        "product_count",
        "delivery_days",
        "delivery_delay_days",
        "review_score",
    ]
    corr = d[numeric_cols].corr(method="spearman")
    corr.to_csv(TABLE_DIR / "driver_spearman_matrix.csv")

    rho, p_value = spearmanr(d["delivery_delay_days"], d["review_score"], nan_policy="omit")
    late_reviews = d.loc[d["delivered_late"], "review_score"].dropna()
    ontime_reviews = d.loc[~d["delivered_late"], "review_score"].dropna()
    stat, u_p = mannwhitneyu(late_reviews, ontime_reviews, alternative="two-sided")

    summary = {
        "spearman_delay_review_rho": float(rho),
        "spearman_delay_review_p_value": float(p_value),
        "mann_whitney_u": float(stat),
        "mann_whitney_p_value": float(u_p),
        "late_review_mean": float(late_reviews.mean()),
        "ontime_review_mean": float(ontime_reviews.mean()),
        "late_review_n": int(len(late_reviews)),
        "ontime_review_n": int(len(ontime_reviews)),
    }
    pd.DataFrame([summary]).to_csv(TABLE_DIR / "driver_test_summary.csv", index=False)
    return corr, summary


def analyze_payment(data: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    p = (
        data["payments"].groupby("payment_type", as_index=False)
        .agg(
            payment_value=("payment_value", "sum"),
            payment_records=("order_id", "count"),
            unique_orders=("order_id", "nunique"),
        )
    )
    p["value_share"] = p["payment_value"] / p["payment_value"].sum()
    p = p.sort_values("payment_value", ascending=False)
    p.to_csv(TABLE_DIR / "payment_analysis.csv", index=False)
    return p


def create_risk_opportunity_tables(
    category: pd.DataFrame,
    region: pd.DataFrame,
    delivery: pd.DataFrame,
    customer: pd.DataFrame,
    kpis: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    overall_freight_ratio = float(kpis.loc[kpis["kpi"] == "Freight-to-revenue ratio", "value"].iloc[0])

    risk_rows = []
    # Category freight burden: require enough item observations to reduce tiny-sample noise.
    for _, r in category[category["items"] >= 500].nlargest(10, "freight_ratio").iterrows():
        if r["freight_ratio"] > overall_freight_ratio:
            risk_rows.append({
                "risk_area": "Category freight burden",
                "segment": r["category"],
                "evidence_metric": "freight_ratio",
                "value": r["freight_ratio"],
                "context": f"{int(r['items']):,} items; {fmt_currency(r['revenue'])} revenue",
                "interpretation": "Freight consumes a higher share of revenue than the overall portfolio average.",
            })
    # Regional freight burden: require 500+ orders.
    for _, r in region[region["orders"] >= 500].nlargest(10, "freight_ratio").iterrows():
        if r["freight_ratio"] > overall_freight_ratio:
            risk_rows.append({
                "risk_area": "Regional freight burden",
                "segment": r["customer_state"],
                "evidence_metric": "freight_ratio",
                "value": r["freight_ratio"],
                "context": f"{int(r['orders']):,} orders; {fmt_currency(r['revenue'])} revenue",
                "interpretation": "Freight burden is above the portfolio average in a sufficiently sized market.",
            })
    # Delivery/customer experience risk from actual delay group results.
    late = delivery[delivery["delay_group"].astype(str).str.contains("late")]
    if not late.empty:
        weighted_late_reviews = np.average(
            late["average_review_score"].dropna(),
            weights=late.loc[late["average_review_score"].notna(), "orders"],
        )
        risk_rows.append({
            "risk_area": "Late-delivery customer experience",
            "segment": "Late deliveries",
            "evidence_metric": "average_review_score",
            "value": weighted_late_reviews,
            "context": "Observed late-delivery groups are compared with on-time/early orders.",
            "interpretation": "Late orders have materially lower average review scores in the observed data; this is associative evidence, not proof of causation.",
        })
    # Retention concentration risk.
    repeat_share = float(kpis.loc[kpis["kpi"] == "Repeat customer rate", "value"].iloc[0])
    risk_rows.append({
        "risk_area": "Customer retention",
        "segment": "Repeat customers",
        "evidence_metric": "repeat_customer_rate",
        "value": repeat_share,
        "context": f"{int((customer['customer_type'] == 'Repeat').sum()):,} repeat customers in customer-level analysis",
        "interpretation": "The historical customer base is dominated by one-time customers, so retention is a material area for investigation.",
    })
    risks = pd.DataFrame(risk_rows)
    risks.to_csv(TABLE_DIR / "risk_analysis.csv", index=False)

    opportunity_rows = []
    # High contribution, meaningful volume categories.
    meaningful_cat = category[category["items"] >= 500].copy()
    for _, r in meaningful_cat.sort_values("contribution_proxy", ascending=False).head(8).iterrows():
        opportunity_rows.append({
            "opportunity_area": "Category growth / mix",
            "segment": r["category"],
            "evidence_metric": "contribution_proxy",
            "value": r["contribution_proxy"],
            "supporting_metric": r["margin_proxy"],
            "interpretation": "Meaningful category contribution proxy with sufficient observed volume.",
        })
    # High-volume regional markets with freight ratio below portfolio average.
    meaningful_regions = region[region["orders"] >= 1000].copy()
    for _, r in meaningful_regions[meaningful_regions["freight_ratio"] < overall_freight_ratio].sort_values(
        "contribution_proxy", ascending=False
    ).head(8).iterrows():
        opportunity_rows.append({
            "opportunity_area": "Regional growth / mix",
            "segment": r["customer_state"],
            "evidence_metric": "contribution_proxy",
            "value": r["contribution_proxy"],
            "supporting_metric": r["freight_ratio"],
            "interpretation": "High-volume market with below-average freight burden.",
        })
    # Repeat customers as a retention opportunity area.
    repeat = customer[customer["customer_type"] == "Repeat"]
    one = customer[customer["customer_type"] == "One-time"]
    opportunity_rows.append({
        "opportunity_area": "Retention",
        "segment": "Repeat customers",
        "evidence_metric": "revenue_per_customer",
        "value": repeat["revenue"].sum() / len(repeat) if len(repeat) else np.nan,
        "supporting_metric": one["revenue"].sum() / len(one) if len(one) else np.nan,
        "interpretation": "Repeat customers have observed aggregate spend per customer above one-time customers because they place multiple orders; retention analysis can target conversion of one-time buyers.",
    })
    opportunities = pd.DataFrame(opportunity_rows)
    opportunities.to_csv(TABLE_DIR / "opportunity_analysis.csv", index=False)
    return risks, opportunities


# -----------------------------------------------------------------------------
# Visualizations
# -----------------------------------------------------------------------------

def create_visualizations(master: pd.DataFrame, category: pd.DataFrame, region: pd.DataFrame, monthly: pd.DataFrame, delivery: pd.DataFrame, customer: pd.DataFrame) -> Dict[str, Path]:
    paths: Dict[str, Path] = {}

    # 1 Monthly revenue
    fig, ax = plt.subplots(figsize=(10, 4.8))
    m = monthly.copy()
    ax.plot(m["purchase_month"], m["revenue"], marker="o", linewidth=1.8)
    ax.set_title("Monthly Item Revenue")
    ax.set_xlabel("Purchase Month")
    ax.set_ylabel("Revenue (BRL)")
    ax.tick_params(axis="x", rotation=45)
    paths["monthly_revenue"] = savefig(fig, "01_monthly_revenue.png")

    # 2 Monthly orders
    fig, ax = plt.subplots(figsize=(10, 4.8))
    ax.plot(m["purchase_month"], m["orders"], marker="o", linewidth=1.8)
    ax.set_title("Monthly Order Volume")
    ax.set_xlabel("Purchase Month")
    ax.set_ylabel("Orders")
    ax.tick_params(axis="x", rotation=45)
    paths["monthly_orders"] = savefig(fig, "02_monthly_orders.png")

    # 3 Category revenue
    top = category.nlargest(12, "revenue").sort_values("revenue")
    fig, ax = plt.subplots(figsize=(9, 6))
    ax.barh(top["category"], top["revenue"])
    ax.set_title("Top Product Categories by Revenue")
    ax.set_xlabel("Revenue (BRL)")
    ax.set_ylabel("Category")
    paths["category_revenue"] = savefig(fig, "03_category_revenue.png")

    # 4 Category contribution proxy
    top = category.nlargest(12, "contribution_proxy").sort_values("contribution_proxy")
    fig, ax = plt.subplots(figsize=(9, 6))
    ax.barh(top["category"], top["contribution_proxy"])
    ax.set_title("Top Product Categories by Freight-Adjusted Contribution Proxy")
    ax.set_xlabel("Contribution Proxy (BRL)")
    ax.set_ylabel("Category")
    paths["category_contribution"] = savefig(fig, "04_category_contribution.png")

    # 5 Category freight ratio
    top = category[category["items"] >= 500].nlargest(12, "freight_ratio").sort_values("freight_ratio")
    fig, ax = plt.subplots(figsize=(9, 6))
    ax.barh(top["category"], top["freight_ratio"] * 100)
    ax.set_title("Categories with Higher Freight-to-Revenue Ratios")
    ax.set_xlabel("Freight / Revenue (%)")
    ax.set_ylabel("Category")
    paths["category_freight"] = savefig(fig, "05_category_freight_ratio.png")

    # 6 State revenue
    top = region.nlargest(12, "revenue").sort_values("revenue")
    fig, ax = plt.subplots(figsize=(9, 6))
    ax.barh(top["customer_state"], top["revenue"])
    ax.set_title("Top Customer States by Revenue")
    ax.set_xlabel("Revenue (BRL)")
    ax.set_ylabel("State")
    paths["state_revenue"] = savefig(fig, "06_state_revenue.png")

    # 7 State freight ratio
    top = region[region["orders"] >= 500].nlargest(12, "freight_ratio").sort_values("freight_ratio")
    fig, ax = plt.subplots(figsize=(9, 6))
    ax.barh(top["customer_state"], top["freight_ratio"] * 100)
    ax.set_title("Freight-to-Revenue Ratio by State (500+ Orders)")
    ax.set_xlabel("Freight / Revenue (%)")
    ax.set_ylabel("State")
    paths["state_freight"] = savefig(fig, "07_state_freight_ratio.png")

    # 8 Delivery delay distribution
    d = master[(master["order_status"] == "delivered") & master["delivery_delay_days"].notna()].copy()
    fig, ax = plt.subplots(figsize=(9, 5))
    ax.hist(d["delivery_delay_days"].clip(-40, 60), bins=40)
    ax.set_title("Distribution of Delivery Delay vs Estimated Date")
    ax.set_xlabel("Delay (days; negative = early)")
    ax.set_ylabel("Orders")
    paths["delay_distribution"] = savefig(fig, "08_delivery_delay_distribution.png")

    # 9 Delay vs review score
    dv = delivery.copy()
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.bar(dv["delay_group"].astype(str), dv["average_review_score"])
    ax.set_title("Average Review Score by Delivery Delay Group")
    ax.set_xlabel("Delivery Group")
    ax.set_ylabel("Average Review Score")
    paths["delay_review"] = savefig(fig, "09_delay_vs_review.png")

    # 10 Repeat vs one-time
    cust_counts = customer["customer_type"].value_counts()
    fig, ax = plt.subplots(figsize=(7, 5))
    ax.bar(cust_counts.index, cust_counts.values)
    ax.set_title("Customer Base: Repeat vs One-Time")
    ax.set_xlabel("Customer Type")
    ax.set_ylabel("Unique Customers")
    paths["customer_type"] = savefig(fig, "10_customer_type.png")

    # 11 Order value distribution
    fin = master[master["total_order_value"].notna()].copy()
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.hist(fin["total_order_value"].clip(0, fin["total_order_value"].quantile(0.99)), bins=50)
    ax.set_title("Order Value Distribution (99th Percentile Capped for Display)")
    ax.set_xlabel("Total Order Value (BRL)")
    ax.set_ylabel("Orders")
    paths["order_value"] = savefig(fig, "11_order_value_distribution.png")

    # 12 Contribution distribution
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.hist(fin["contribution_proxy"].clip(fin["contribution_proxy"].quantile(0.01), fin["contribution_proxy"].quantile(0.99)), bins=50)
    ax.set_title("Freight-Adjusted Contribution Proxy Distribution")
    ax.set_xlabel("Contribution Proxy (BRL)")
    ax.set_ylabel("Orders")
    paths["contribution_distribution"] = savefig(fig, "12_contribution_distribution.png")

    return paths


def create_dashboard_image(kpis: pd.DataFrame, category: pd.DataFrame, region: pd.DataFrame, monthly: pd.DataFrame, delivery: pd.DataFrame) -> Path:
    # Dashboard mock-up using matplotlib; designed for the internship report.
    fig = plt.figure(figsize=(14, 9))
    ax = fig.add_axes([0.02, 0.02, 0.96, 0.96])
    ax.axis("off")
    ax.text(0.01, 0.97, "Executive E-Commerce Analytics Dashboard", fontsize=20, fontweight="bold", va="top")
    ax.text(0.01, 0.93, "Olist historical dataset | 2016-2018 | Freight-adjusted contribution proxy", fontsize=10, va="top")

    lookup = {r.kpi: r.value for r in kpis.itertuples()}
    cards = [
        ("Revenue", fmt_currency(lookup["Total item revenue"])),
        ("Orders", fmt_num(lookup["Total registered orders"])),
        ("Customers", fmt_num(lookup["Unique customers"])),
        ("AOV", fmt_currency(lookup["Average order value incl. freight"])),
        ("Freight Ratio", fmt_pct(lookup["Freight-to-revenue ratio"])),
        ("Contribution Margin Proxy", fmt_pct(lookup["Contribution proxy margin"])),
    ]
    x0, y0 = 0.01, 0.80
    w, h = 0.145, 0.10
    gap = 0.015
    for i, (label, value) in enumerate(cards):
        x = x0 + i * (w + gap)
        rect = plt.Rectangle((x, y0), w, h, fill=False, linewidth=1.5, transform=ax.transAxes)
        ax.add_patch(rect)
        ax.text(x + 0.012, y0 + 0.065, label, fontsize=9, transform=ax.transAxes)
        ax.text(x + 0.012, y0 + 0.025, value, fontsize=12, fontweight="bold", transform=ax.transAxes)

    # Small charts in dashboard.
    panels = [
        (0.01, 0.42, 0.30, 0.32),
        (0.35, 0.42, 0.30, 0.32),
        (0.69, 0.42, 0.30, 0.32),
        (0.01, 0.07, 0.48, 0.28),
        (0.52, 0.07, 0.47, 0.28),
    ]
    # 1 monthly revenue
    a1 = fig.add_axes(panels[0])
    a1.plot(monthly["purchase_month"], monthly["revenue"], linewidth=1.4)
    a1.set_title("Revenue Trend", fontsize=10)
    a1.tick_params(axis="x", labelrotation=45, labelsize=7)
    a1.tick_params(axis="y", labelsize=7)

    # 2 category contribution
    a2 = fig.add_axes(panels[1])
    c = category.nlargest(6, "contribution_proxy").sort_values("contribution_proxy")
    a2.barh(c["category"], c["contribution_proxy"])
    a2.set_title("Contribution Proxy by Category", fontsize=10)
    a2.tick_params(labelsize=7)

    # 3 state freight
    a3 = fig.add_axes(panels[2])
    s = region[region["orders"] >= 500].nlargest(6, "freight_ratio").sort_values("freight_ratio")
    a3.barh(s["customer_state"], s["freight_ratio"] * 100)
    a3.set_title("Freight Ratio: Higher-Load States", fontsize=10)
    a3.tick_params(labelsize=8)

    # 4 delivery review
    a4 = fig.add_axes(panels[3])
    a4.bar(delivery["delay_group"].astype(str), delivery["average_review_score"])
    a4.set_title("Delivery Delay vs Review Score", fontsize=10)
    a4.tick_params(axis="x", labelrotation=15, labelsize=8)
    a4.tick_params(axis="y", labelsize=8)

    # 5 executive callouts
    a5 = fig.add_axes(panels[4]); a5.axis("off")
    a5.text(0.02, 0.88, "Executive attention areas", fontsize=11, fontweight="bold")
    top_cat = category.nlargest(1, "contribution_proxy").iloc[0]
    top_rev_cat = category.nlargest(1, "revenue").iloc[0]
    top_state = region[region["orders"] >= 500].nlargest(1, "freight_ratio").iloc[0]
    a5.text(0.03, 0.67, f"Highest contribution proxy category: {top_cat['category']}", fontsize=9)
    a5.text(0.03, 0.50, f"Highest-revenue category: {top_rev_cat['category']}", fontsize=9)
    a5.text(0.03, 0.33, f"Highest freight ratio among 500+ order states: {top_state['customer_state']}", fontsize=9)
    a5.text(0.03, 0.16, "Interpretation: contribution proxy is freight-adjusted, not full profit.", fontsize=8)

    path = FIG_DIR / "00_executive_dashboard.png"
    fig.savefig(path, dpi=170, bbox_inches="tight")
    plt.close(fig)
    return path


# -----------------------------------------------------------------------------
# Executive summary and report
# -----------------------------------------------------------------------------

def build_summary_text(
    data: Dict[str, pd.DataFrame],
    master: pd.DataFrame,
    kpis: pd.DataFrame,
    category: pd.DataFrame,
    region: pd.DataFrame,
    delivery: pd.DataFrame,
    driver_tests: Dict[str, float],
) -> str:
    lookup = {r.kpi: r.value for r in kpis.itertuples()}
    top_rev_cat = category.nlargest(1, "revenue").iloc[0]
    top_contrib_cat = category.nlargest(1, "contribution_proxy").iloc[0]
    top_state = region.nlargest(1, "revenue").iloc[0]
    freight_state = region[region["orders"] >= 500].nlargest(1, "freight_ratio").iloc[0]
    late = delivery[delivery["delay_group"].astype(str).str.contains("late")]
    late_n = int(late["orders"].sum()) if not late.empty else 0

    lines = [
        "# Executive Analytics Summary",
        "",
        "## Dataset",
        f"- Olist orders table: {len(data['orders']):,} orders.",
        f"- Item table: {len(data['items']):,} order-item records.",
        f"- Purchase period: {data['orders']['order_purchase_timestamp'].min().date()} to {data['orders']['order_purchase_timestamp'].max().date()}.",
        "",
        "## Financial metric definition",
        "The dataset does not provide complete product/operating cost information. The project therefore uses a **Freight-Adjusted Contribution Proxy = item revenue - freight value**. This is not accounting profit.",
        "",
        "## Actual KPIs",
        f"- Total registered orders: {int(lookup['Total registered orders']):,}",
        f"- Delivered orders: {int(lookup['Delivered orders']):,} ({fmt_pct(lookup['Delivered orders'] / lookup['Total registered orders']) if lookup['Total registered orders'] else 'N/A'})",
        f"- Item revenue: {fmt_currency(lookup['Total item revenue'])}",
        f"- Freight: {fmt_currency(lookup['Total freight'])}",
        f"- Contribution proxy: {fmt_currency(lookup['Freight-adjusted contribution proxy'])}",
        f"- Freight-to-revenue ratio: {fmt_pct(lookup['Freight-to-revenue ratio'])}",
        f"- Average order value incl. freight: {fmt_currency(lookup['Average order value incl. freight'])}",
        f"- Average delivery time: {lookup['Average delivery days']:.2f} days",
        f"- Late delivery rate: {fmt_pct(lookup['Late delivery rate'])}",
        f"- Average review score: {lookup['Average review score']:.2f}",
        f"- Repeat customer rate: {fmt_pct(lookup['Repeat customer rate'])}",
        "",
        "## Evidence-based findings",
        f"1. **Revenue concentration:** {top_rev_cat['category']} is the highest-revenue product category at {fmt_currency(top_rev_cat['revenue'])}.",
        f"2. **Freight-adjusted contribution:** {top_contrib_cat['category']} has the largest contribution proxy at {fmt_currency(top_contrib_cat['contribution_proxy'])}, with a freight ratio of {fmt_pct(top_contrib_cat['freight_ratio'])}.",
        f"3. **Regional scale:** {top_state['customer_state']} generates the largest state-level revenue at {fmt_currency(top_state['revenue'])} across {int(top_state['orders']):,} orders.",
        f"4. **Regional freight burden:** Among states with at least 500 orders, {freight_state['customer_state']} has the highest freight-to-revenue ratio at {fmt_pct(freight_state['freight_ratio'])}.",
        f"5. **Customer experience:** {late_n:,} delivered orders were late in the delay analysis; delivery delay and review score show a Spearman association of {driver_tests['spearman_delay_review_rho']:.3f} (p < 0.001).",
        "",
        "## Customer retention finding",
        f"Only {fmt_pct(lookup['Repeat customer rate'])} of unique customers are repeat customers in the historical data, while repeat-customer orders represent {fmt_pct(lookup['Order share from repeat customers'])} of all orders. This supports investigating retention and second-purchase conversion rather than assuming high customer lifetime value.",
        "",
        "## Interpretation and limitations",
        "- The financial proxy excludes COGS and other operating costs, so it must not be interpreted as true profit.",
        "- High freight ratio is evidence of freight burden, not proof that freight alone causes weak profitability.",
        "- The delay/review relationship is statistical association, not proof of causation.",
        "- Very small states/categories should be interpreted cautiously; the project uses sample-size thresholds for operational risk flags.",
        "- The dataset is historical (2016-2018) and may not represent current e-commerce operations.",
    ]
    text = "\n".join(lines)
    (OUTPUT_DIR / "executive_summary.md").write_text(text, encoding="utf-8")
    return text


def create_report(
    data: Dict[str, pd.DataFrame],
    master: pd.DataFrame,
    kpis: pd.DataFrame,
    category: pd.DataFrame,
    region: pd.DataFrame,
    customer: pd.DataFrame,
    monthly: pd.DataFrame,
    delivery: pd.DataFrame,
    risks: pd.DataFrame,
    opportunities: pd.DataFrame,
    driver_tests: Dict[str, float],
    fig_paths: Dict[str, Path],
    dashboard_path: Path,
) -> None:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    styles = doc.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(10.5)
    for style_name in ["Title", "Heading 1", "Heading 2", "Heading 3"]:
        styles[style_name].font.name = "Aptos Display"

    # Cover
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.space_after = Pt(20)
    r = p.add_run("GLOBAL E-COMMERCE SALES &\nPROFITABILITY OPTIMIZATION ANALYTICS")
    r.bold = True
    r.font.size = Pt(22)
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run("IBM Data Analytics Internship Project").bold = True
    doc.add_paragraph("")
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.add_run("Brazilian E-Commerce Public Dataset by Olist")
    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.add_run(f"Analysis period: {data['orders']['order_purchase_timestamp'].min().date()} to {data['orders']['order_purchase_timestamp'].max().date()}")
    doc.add_paragraph("")
    p5 = doc.add_paragraph()
    p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p5.add_run("Prepared from the actual uploaded dataset; no synthetic records were used.").italic = True
    doc.add_page_break()

    # Executive summary
    doc.add_heading("1. Executive Summary", level=1)
    lookup = {r.kpi: r.value for r in kpis.itertuples()}
    paragraphs = [
        f"The project analyzes the Olist Brazilian e-commerce dataset to understand sales performance, freight burden, category mix, regional performance, customer repeat behavior, and delivery-related customer experience. The dataset contains {len(data['orders']):,} orders and {len(data['items']):,} order-item records.",
        f"Across orders with item-level financial records, item revenue is {fmt_currency(lookup['Total item revenue'])} and freight is {fmt_currency(lookup['Total freight'])}, producing a freight-to-revenue ratio of {fmt_pct(lookup['Freight-to-revenue ratio'])}. The project reports a freight-adjusted contribution proxy of {fmt_currency(lookup['Freight-adjusted contribution proxy'])}.",
        f"Operationally, {int(lookup['Delivered orders']):,} orders are recorded as delivered. Among delivered orders with valid timing fields, average delivery time is {lookup['Average delivery days']:.2f} days and the late-delivery rate is {fmt_pct(lookup['Late delivery rate'])}.",
        f"Customer retention is a significant analytical area: the repeat-customer rate is {fmt_pct(lookup['Repeat customer rate'])}, while repeat-customer orders represent {fmt_pct(lookup['Order share from repeat customers'])} of all orders. The analysis also finds a negative association between delivery delay and review score (Spearman rho {driver_tests['spearman_delay_review_rho']:.3f}; p < 0.001).",
    ]
    for t in paragraphs:
        doc.add_paragraph(t)
    doc.add_paragraph("Important limitation: Olist does not provide complete accounting cost information. Therefore, the contribution proxy is not true accounting profit.").runs[0].bold = True
    add_figure(doc, dashboard_path, "Figure 1. Executive dashboard view generated from the actual Olist data.")

    # Introduction / business problem
    doc.add_heading("2. Introduction", level=1)
    doc.add_paragraph("E-commerce analytics requires more than reporting sales volume. The project follows a business-intelligence sequence: KPIs describe what is happening; trends show direction; driver analysis investigates measurable associations; risk and opportunity analysis identifies areas for management attention; and recommendations connect evidence to action.")

    doc.add_heading("3. Business Problem", level=1)
    doc.add_paragraph("Identify which product categories, customer segments, and geographic regions contribute strongly to sales and freight-adjusted contribution while identifying freight burden, delivery problems, and customer-experience risks.")

    doc.add_heading("4. Business Objective", level=1)
    doc.add_paragraph("Improve e-commerce resource efficiency and commercial decision-making using evidence from sales, logistics, customer, payment, and review data.")

    doc.add_heading("5. Analytics Objectives", level=1)
    for q in [
        "Measure category-level revenue and freight-adjusted contribution.",
        "Quantify freight burden across customer states.",
        "Measure delivery duration and late delivery prevalence.",
        "Examine the association between delivery delay and review scores.",
        "Measure repeat-customer behavior and customer value.",
        "Identify evidence-based risks, opportunities, and business actions.",
    ]:
        doc.add_paragraph(q, style="List Bullet")

    # Dataset
    doc.add_heading("6. Dataset Description", level=1)
    doc.add_paragraph(f"Source: {DATASET_URL}")
    raw_inventory = pd.read_csv(TABLE_DIR / "raw_table_inventory.csv")
    raw_geo = int(raw_inventory.loc[raw_inventory["table"] == "geolocation", "rows"].iloc[0])
    cleaned_geo = len(data["geolocation"])
    doc.add_paragraph(f"The uploaded dataset contains {len(data['orders']):,} order records, {len(data['items']):,} order-item records, {len(data['customers']):,} customer records, {len(data['products']):,} product records, {len(data['payments']):,} payment records, {len(data['reviews']):,} review records, {len(data['sellers']):,} seller records, {raw_geo:,} raw geolocation rows ({cleaned_geo:,} after exact-duplicate cleaning), and {len(data['translation']):,} category translation rows.")
    doc.add_paragraph(f"Order purchase timestamps span {data['orders']['order_purchase_timestamp'].min().date()} to {data['orders']['order_purchase_timestamp'].max().date()}.")
    inventory = pd.read_csv(TABLE_DIR / "raw_table_inventory.csv")
    add_table(doc, inventory[["table", "rows", "columns", "exact_duplicate_rows", "missing_cells"]], "Table 1. Raw Table Inventory", max_rows=12)

    # Methodology
    doc.add_heading("7. Data Preparation and Methodology", level=1)
    for t in [
        "Exact duplicate rows were removed without collapsing legitimate multi-item orders.",
        "Dates were converted using pandas datetime parsing with invalid strings becoming missing values.",
        "Financial fields were coerced to numeric types; item rows missing price or freight were excluded from contribution calculations.",
        "Missing product categories were treated as an explicit unknown category rather than fabricated labels.",
        "Payment data were aggregated to order level and the dominant payment type was selected by recorded payment value.",
        "Review data were aggregated to order level using the mean available review score; missing reviews were not imputed.",
        "The order master was built with one row per Olist order. Item-level category analysis was retained separately to avoid double-counting when multi-item orders contain several categories.",
        "Financial KPIs are based on orders with item-level records because price and freight are stored at item level.",
    ]:
        doc.add_paragraph(t, style="List Bullet")

    doc.add_heading("8. Data Dictionary", level=1)
    dictionary = pd.read_csv(TABLE_DIR / "data_dictionary.csv")
    add_table(doc, dictionary[["column_name", "data_type", "description", "missing_percentage", "analytical_role"]], "Table 2. Analytical Data Dictionary", max_rows=22)

    # KPIs
    doc.add_heading("9. KPI Analysis", level=1)
    kpi_view = kpis.copy()
    kpi_view["display_value"] = [
        fmt_num(v) if u == "count" else fmt_currency(v) if u == "BRL" else fmt_pct(v) if u == "ratio" else f"{v:.2f}"
        for v, u in zip(kpi_view["value"], kpi_view["unit"])
    ]
    add_table(doc, kpi_view[["kpi", "display_value", "definition"]], "Table 3. Key Performance Indicators", max_rows=20)

    # EDA
    doc.add_heading("10. Exploratory Data Analysis", level=1)
    doc.add_paragraph("The exploratory stage examines overall sales distribution, category concentration, geographic performance, customer type, delivery delay, payment mix, and monthly trends. The generated charts are included below and are based on the actual uploaded dataset.")
    add_figure(doc, fig_paths["monthly_revenue"], "Figure 2. Monthly item revenue trend.")
    add_figure(doc, fig_paths["category_revenue"], "Figure 3. Highest-revenue product categories.")
    add_figure(doc, fig_paths["state_revenue"], "Figure 4. Highest-revenue customer states.")

    # Category
    doc.add_heading("11. Product Category Analysis", level=1)
    top_cat = category.nlargest(8, "revenue").copy()
    top_cat["Revenue (BRL)"] = top_cat["revenue"].map(lambda x: f"{x:,.0f}")
    top_cat["Contribution Proxy (BRL)"] = top_cat["contribution_proxy"].map(lambda x: f"{x:,.0f}")
    top_cat["Freight Ratio"] = (top_cat["freight_ratio"] * 100).map(lambda x: f"{x:.2f}%")
    top_cat["Margin Proxy"] = (top_cat["margin_proxy"] * 100).map(lambda x: f"{x:.2f}%")
    add_table(doc, top_cat[["category", "items", "orders", "Revenue (BRL)", "Contribution Proxy (BRL)", "Freight Ratio", "Margin Proxy"]], "Table 4. Top Categories by Revenue", max_rows=8)
    add_figure(doc, fig_paths["category_contribution"], "Figure 5. Categories with the largest freight-adjusted contribution proxy.")
    add_figure(doc, fig_paths["category_freight"], "Figure 6. Categories with higher freight-to-revenue ratios among categories with 500+ items.")

    # Regional
    doc.add_heading("12. Regional Analysis", level=1)
    region_view = region.nlargest(12, "revenue").copy()
    region_view["Revenue (BRL)"] = region_view["revenue"].map(lambda x: f"{x:,.0f}")
    region_view["Freight Ratio"] = (region_view["freight_ratio"] * 100).map(lambda x: f"{x:.2f}%")
    region_view["Late Rate"] = (region_view["late_rate"] * 100).map(lambda x: f"{x:.2f}%")
    region_view["Avg Review"] = region_view["avg_review_score"].map(lambda x: f"{x:.2f}" if pd.notna(x) else "N/A")
    add_table(doc, region_view[["customer_state", "orders", "Revenue (BRL)", "Freight Ratio", "Late Rate", "Avg Review"]], "Table 5. Regional Performance", max_rows=12)
    add_figure(doc, fig_paths["state_freight"], "Figure 7. Freight-to-revenue ratio by state for states with at least 500 orders.")

    # Customers
    doc.add_heading("13. Customer Analysis", level=1)
    cust_summary = (
        customer.groupby("customer_type")
        .agg(
            customers=("customer_unique_id", "nunique"),
            orders=("orders", "sum"),
            revenue=("revenue", "sum"),
            contribution_proxy=("contribution_proxy", "sum"),
            revenue_per_customer=("revenue", "mean"),
        )
        .reset_index()
    )
    cust_summary["revenue"] = cust_summary["revenue"].round(2)
    cust_summary["contribution_proxy"] = cust_summary["contribution_proxy"].round(2)
    cust_summary["revenue_per_customer"] = cust_summary["revenue_per_customer"].round(2)
    add_table(doc, cust_summary, "Table 6. Customer Type Summary", max_rows=5)
    add_figure(doc, fig_paths["customer_type"], "Figure 8. Unique customers by repeat/one-time classification.")
    doc.add_paragraph("Repeat classification uses customer_unique_id and the historical number of orders. The analysis does not label the result as customer lifetime value because the dataset covers a fixed historical period and lacks forward-looking retention information.")

    # Delivery
    doc.add_heading("14. Delivery Performance and Customer Satisfaction", level=1)
    delivery_view = delivery.copy()
    delivery_view["average_delivery_days"] = delivery_view["average_delivery_days"].round(2)
    delivery_view["average_delay_days"] = delivery_view["average_delay_days"].round(2)
    delivery_view["average_review_score"] = delivery_view["average_review_score"].round(2)
    delivery_view["order_share"] = (delivery_view["order_share"] * 100).round(2)
    add_table(doc, delivery_view, "Table 7. Delivery Delay Groups", max_rows=5)
    add_figure(doc, fig_paths["delay_distribution"], "Figure 9. Delivery delay distribution relative to estimated date.")
    add_figure(doc, fig_paths["delay_review"], "Figure 10. Average review score across delivery-delay groups.")
    doc.add_paragraph(
        f"Statistical evidence: Spearman correlation between delivery delay days and review score is {driver_tests['spearman_delay_review_rho']:.3f} with p < 0.001. The Mann-Whitney U test comparing review scores for late versus on-time/early orders is also significant (p < 0.001). The direction indicates that larger delays are associated with lower reviews in this dataset; this is association, not causal proof."
    )

    # Trends
    doc.add_heading("15. Trend Analysis", level=1)
    doc.add_paragraph("The time series covers the Olist historical period and is strongest from 2017 through 2018; early and final partial months should not be interpreted as complete comparable periods.")
    add_figure(doc, fig_paths["monthly_orders"], "Figure 11. Monthly order volume trend.")
    peak = monthly.loc[monthly["revenue"].idxmax()]
    doc.add_paragraph(f"The highest-revenue month in the item-level series is {peak['purchase_month']} with {int(peak['orders']):,} orders and {fmt_currency(peak['revenue'])} in item revenue.")

    # Drivers
    doc.add_heading("16. Driver Analysis", level=1)
    doc.add_paragraph("Driver analysis uses descriptive comparisons and non-parametric/statistical association. The analysis deliberately avoids causal language when the dataset is observational.")
    corr = pd.read_csv(TABLE_DIR / "driver_spearman_matrix.csv", index_col=0)
    selected = corr.loc[["freight_ratio", "delivery_delay_days", "review_score"], ["item_revenue", "freight_value", "contribution_proxy", "delivery_delay_days", "review_score"]]
    add_table(doc, selected.reset_index().round(3), "Table 8. Selected Spearman Associations", max_rows=8)

    # Risks / opportunities
    doc.add_page_break()
    doc.add_heading("17. Risk Analysis", level=1)
    risk_view = risks.head(8).copy()
    risk_view["Evidence Value"] = risk_view["value"].map(lambda x: f"{x:.3f}" if pd.notna(x) else "N/A")
    add_table(doc, risk_view[["risk_area", "segment", "evidence_metric", "Evidence Value", "context"]], "Table 9. Data-Driven Risk Signals", max_rows=8)
    for _, r in risk_view.iterrows():
        doc.add_paragraph(f"{r['risk_area']} — {r['segment']}: {r['interpretation']}")

    doc.add_heading("18. Opportunity Analysis", level=1)
    opp_view = opportunities.head(8).copy()
    opp_view["Evidence Value"] = opp_view["value"].map(lambda x: f"{x:,.0f}" if pd.notna(x) else "N/A")
    opp_view["Supporting Value"] = opp_view["supporting_metric"].map(lambda x: f"{x:.3f}" if pd.notna(x) else "N/A")
    add_table(doc, opp_view[["opportunity_area", "segment", "evidence_metric", "Evidence Value", "Supporting Value"]], "Table 10. Data-Driven Opportunities", max_rows=8)
    for _, r in opp_view.iterrows():
        doc.add_paragraph(f"{r['opportunity_area']} — {r['segment']}: {r['interpretation']}")

    # Recommendations
    doc.add_heading("19. Actionable Business Recommendations", level=1)
    top_rev_cat = category.nlargest(1, "revenue").iloc[0]
    top_contrib_cat = category.nlargest(1, "contribution_proxy").iloc[0]
    high_freight_cat = category[category["items"] >= 500].nlargest(1, "freight_ratio").iloc[0]
    high_freight_state = region[region["orders"] >= 500].nlargest(1, "freight_ratio").iloc[0]
    recommendations = [
        (
            "Protect and scale strong category economics",
            f"{top_contrib_cat['category']} has the largest contribution proxy ({fmt_currency(top_contrib_cat['contribution_proxy'])}).",
            "Review assortment depth, inventory availability and cross-sell opportunities around strong-contribution categories while monitoring freight ratio.",
            "Contribution proxy and freight ratio",
        ),
        (
            "Investigate freight burden in high-load categories",
            f"{high_freight_cat['category']} has a freight ratio of {fmt_pct(high_freight_cat['freight_ratio'])} among categories with 500+ items.",
            "Review shipping lanes, package dimensions, seller fulfillment choices and shipping-price rules for high-load categories.",
            "Freight ratio and contribution proxy margin",
        ),
        (
            "Prioritize operational review in high-freight regions",
            f"{high_freight_state['customer_state']} has a freight ratio of {fmt_pct(high_freight_state['freight_ratio'])} among states with 500+ orders.",
            "Evaluate carrier mix, regional fulfillment strategy and shipping thresholds before changing customer-facing pricing.",
            "Regional freight ratio and late-delivery rate",
        ),
        (
            "Use delivery performance as a customer-experience KPI",
            f"Late deliveries show lower average review scores and a negative Spearman association with review score ({driver_tests['spearman_delay_review_rho']:.3f}).",
            "Track late-delivery rate by carrier/region/category and create an operational exception process for chronically delayed segments.",
            "Late-delivery rate and average review score",
        ),
        (
            "Build a second-purchase retention program",
            f"The repeat-customer rate is {fmt_pct(float(kpis.loc[kpis['kpi'] == 'Repeat customer rate', 'value'].iloc[0]))}.",
            "Test post-purchase messaging, category cross-sell and timed second-purchase offers for one-time customers, then measure repeat-order conversion.",
            "Repeat customer rate and second-order conversion",
        ),
    ]
    for title, evidence, action, kpi in recommendations:
        doc.add_heading(title, level=3)
        doc.add_paragraph(f"Evidence: {evidence}")
        doc.add_paragraph(f"Recommended action: {action}")
        doc.add_paragraph(f"KPI to monitor: {kpi}")

    # Dashboard + limitations
    doc.add_heading("20. Dashboard Design", level=1)
    doc.add_paragraph("The dashboard is designed for executives and operating managers who need quick answers about sales scale, freight burden, category performance, customer retention, delivery performance, and evidence-based action areas.")
    add_figure(doc, dashboard_path, "Figure 12. Final executive dashboard design.")

    doc.add_heading("21. Limitations", level=1)
    for t in [
        "The Olist dataset does not contain full accounting cost data; contribution proxy is therefore not true profit.",
        "Historical period ends in 2018, so results should not be treated as current market performance.",
        "Observed associations do not establish causality.",
        "Customer retention is measured from historical repeat ordering and should not be treated as a forward-looking CLV model.",
        "Some regions/categories have low volume; sample-size thresholds are used for risk flags where practical.",
        "Freight value is an observed dataset field; the analysis does not independently verify carrier-level cost allocation.",
    ]:
        doc.add_paragraph(t, style="List Bullet")

    doc.add_heading("22. Conclusion", level=1)
    doc.add_paragraph("The completed analysis converts the Olist transactional data into a reproducible Business Intelligence workflow. The project combines sales, freight, category, regional, customer and delivery evidence, and translates the observed patterns into measurable management actions. The most important analytical discipline is preserving the distinction between actual dataset evidence and unsupported assumptions, especially for profitability and causal explanations.")

    doc.add_heading("23. References", level=1)
    doc.add_paragraph(f"Olist / Kaggle: {DATASET_URL}")
    doc.add_paragraph("Python libraries: pandas, NumPy, Matplotlib, SciPy, python-docx.")

    doc.add_heading("24. Appendix — Additional Output Files", level=1)
    doc.add_paragraph("The project also creates detailed CSV outputs in outputs/tables/ and report-ready figures in outputs/figures/.")

    doc.save(REPORT_PATH)


# -----------------------------------------------------------------------------
# README / support files
# -----------------------------------------------------------------------------

def create_support_files() -> None:
    readme = f"""# Global E-Commerce Sales & Profitability Optimization Analytics

IBM Data Analytics Internship Project

## Overview

This project analyzes the **Brazilian E-Commerce Public Dataset by Olist** using a reproducible Python pipeline. It covers KPI analysis, trends, category performance, regional freight burden, customer repeat behavior, delivery performance, statistical association, risk/opportunity evidence, and business recommendations.

## Dataset

Source: {DATASET_URL}

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
.venv\\Scripts\\activate
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

"""
    (PROJECT_ROOT / "README.md").write_text(readme, encoding="utf-8")

    req = """pandas==2.2.3
numpy==2.1.3
matplotlib==3.9.2
scipy==1.14.1
python-docx==1.1.2
"""
    (PROJECT_ROOT / "requirements.txt").write_text(req, encoding="utf-8")

    (PROJECT_ROOT / ".gitignore").write_text(
        """.venv/\n__pycache__/\n*.pyc\n.DS_Store\nThumbs.db\n# Keep raw Olist data out of GitHub unless explicitly allowed\ndata/raw/*.csv\n""",
        encoding="utf-8",
    )
    data_readme = f"""# Raw Dataset\n\nDownload the Brazilian E-Commerce Public Dataset by Olist from:\n\n{DATASET_URL}\n\nExtract the CSV files into this folder before running `python project.py`.\n"""
    (RAW_DIR / "README.md").write_text(data_readme, encoding="utf-8")


# -----------------------------------------------------------------------------
# Main pipeline
# -----------------------------------------------------------------------------

def main() -> None:
    ensure_dirs()
    create_support_files()

    print("PHASE 1 — INSPECTING THE ACTUAL OLIST DATASET")
    data = load_data()
    raw_inventory = inspect_raw_data(data)
    print(raw_inventory[["table", "rows", "columns"]].to_string(index=False))

    print("PHASE 2 — CLEANING AND VALIDATION")
    data = clean_data(data)

    print("PHASE 3 — BUILDING ANALYTICAL MASTER TABLE")
    master = build_analytical_master(data)
    item_category = build_item_category_table(data)
    create_data_dictionary(data, master)

    # Validate join grain.
    if master["order_id"].nunique() != len(master):
        raise ValueError("Order-level master table is not one row per order.")

    print("PHASE 4 — KPI / EDA / DRIVER ANALYSIS")
    kpis = calculate_kpis(master)
    category = analyze_category(item_category)
    region = analyze_region(master)
    customer = analyze_customer(master)
    monthly = analyze_monthly(master)
    delivery = analyze_delivery(master)
    _, driver_tests = analyze_drivers(master)
    analyze_payment(data)
    risks, opportunities = create_risk_opportunity_tables(category, region, delivery, customer, kpis)

    print("PHASE 5 — CREATING VISUALIZATIONS")
    fig_paths = create_visualizations(master, category, region, monthly, delivery, customer)
    dashboard = create_dashboard_image(kpis, category, region, monthly, delivery)
    build_summary_text(data, master, kpis, category, region, delivery, driver_tests)

    print("PHASE 6 — GENERATING FINAL REPORT")
    create_report(data, master, kpis, category, region, customer, monthly, delivery, risks, opportunities, driver_tests, fig_paths, dashboard)

    print("PHASE 7 — VALIDATION")
    # Lightweight output validation.
    expected = [
        OUTPUT_DIR / "analytical_master_table.csv",
        OUTPUT_DIR / "executive_summary.md",
        TABLE_DIR / "kpis.csv",
        TABLE_DIR / "data_dictionary.csv",
        TABLE_DIR / "category_analysis.csv",
        TABLE_DIR / "regional_analysis.csv",
        TABLE_DIR / "customer_analysis.csv",
        TABLE_DIR / "delivery_analysis.csv",
        TABLE_DIR / "risk_analysis.csv",
        TABLE_DIR / "opportunity_analysis.csv",
        REPORT_PATH,
        dashboard,
    ]
    missing = [str(p) for p in expected if not p.exists() or p.stat().st_size == 0]
    if missing:
        raise RuntimeError(f"Output validation failed; missing/empty outputs: {missing}")

    lookup = {r.kpi: r.value for r in kpis.itertuples()}
    print("\nPROJECT COMPLETED AND VALIDATED")
    print("=" * 75)
    print(f"Orders: {int(lookup['Total registered orders']):,}")
    print(f"Item revenue: {fmt_currency(lookup['Total item revenue'])}")
    print(f"Freight ratio: {fmt_pct(lookup['Freight-to-revenue ratio'])}")
    print(f"Contribution proxy: {fmt_currency(lookup['Freight-adjusted contribution proxy'])}")
    print(f"Late delivery rate: {fmt_pct(lookup['Late delivery rate'])}")
    print(f"Repeat customer rate: {fmt_pct(lookup['Repeat customer rate'])}")
    print(f"Report: {REPORT_PATH}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nPROJECT ERROR: {type(exc).__name__}: {exc}")
        raise
