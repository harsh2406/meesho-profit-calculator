from __future__ import annotations

import hashlib
import io
import re
from typing import Any

import pandas as pd
import streamlit as st


APP_TITLE = "Meesho Seller Profit Calculator"
REQUIRED_COLUMNS = {
    "supplier_sku": "Supplier SKU",
    "live_order_status": "Live Order Status",
    "final_settlement_amount": "Final Settlement Amount",
}
OPTIONAL_COLUMNS = {
    "quantity": "Quantity",
    "sub_order_no": "Sub Order No",
    "order_date": "Order Date",
    "product_name": "Product Name",
}
VALID_STATUSES = ("Delivered", "Return", "RTO")


st.set_page_config(
    page_title=APP_TITLE,
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ---------- Helpers ----------
def normalize_header(value: Any) -> str:
    """Normalize a column label so minor spaces/punctuation differences do not matter."""
    return re.sub(r"[^a-z0-9]", "", str(value).strip().lower())


def find_column(df: pd.DataFrame, requested_name: str) -> str | None:
    wanted = normalize_header(requested_name)
    for column in df.columns:
        if normalize_header(column) == wanted:
            return str(column)
    return None


def clean_sku(value: Any) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def parse_amount(value: Any) -> float:
    """Convert Excel currency/number cells to a float. Blank or invalid values become 0."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return 0.0

    is_parenthesized_negative = text.startswith("(") and text.endswith(")")
    text = text.replace(",", "").replace("₹", "").replace("INR", "")
    text = text.replace("(", "").replace(")", "").strip()
    text = re.sub(r"[^0-9.\-]", "", text)
    try:
        amount = float(text)
        return -abs(amount) if is_parenthesized_negative else amount
    except ValueError:
        return 0.0


def classify_status(value: Any) -> str | None:
    if pd.isna(value):
        return None
    status = str(value).strip().lower()
    if not status:
        return None
    if "rto" in status:
        return "RTO"
    if "return" in status:
        return "Return"
    if "deliver" in status:
        return "Delivered"
    # Preserve nonblank, unrecognized statuses so the app can report them.
    return str(value).strip()


def load_order_payments(file_bytes: bytes) -> tuple[pd.DataFrame, dict[str, str], list[str]]:
    """Load Order Payments, finding the real header row and required columns."""
    excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
    if "Order Payments" not in excel_file.sheet_names:
        available = ", ".join(excel_file.sheet_names)
        raise ValueError(
            "The workbook does not contain a sheet named 'Order Payments'. "
            f"Sheets found: {available or 'none'}"
        )

    # Meesho's payment workbook uses a grouped heading on row 1 and the actual
    # column names on row 2. Try row 2 first, then row 1 for other export formats.
    raw_df = None
    column_map: dict[str, str] = {}
    missing_required: list[str] = []

    for header_row in (1, 0, 2):
        candidate = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name="Order Payments",
            header=header_row,
        )
        candidate_map: dict[str, str] = {}
        for key, display_name in REQUIRED_COLUMNS.items():
            found = find_column(candidate, display_name)
            if found is not None:
                candidate_map[key] = found

        if len(candidate_map) == len(REQUIRED_COLUMNS):
            raw_df = candidate.copy()
            column_map = candidate_map
            break

        if len(candidate_map) > len(missing_required):
            missing_required = [
                display_name
                for key, display_name in REQUIRED_COLUMNS.items()
                if key not in candidate_map
            ]

    if raw_df is None:
        raise ValueError(
            "Could not find all required columns in the 'Order Payments' sheet. "
            "Required: Supplier SKU, Live Order Status, and Final Settlement Amount. "
            f"Missing: {', '.join(missing_required) or 'unknown'}"
        )

    for key, display_name in OPTIONAL_COLUMNS.items():
        found = find_column(raw_df, display_name)
        if found is not None:
            column_map[key] = found

    # Convert the source columns to a consistent internal format.
    df = pd.DataFrame()
    df["Supplier SKU"] = raw_df[column_map["supplier_sku"]].map(clean_sku)
    df["Live Order Status"] = raw_df[column_map["live_order_status"]].map(classify_status)
    df["Final Settlement Amount"] = raw_df[
        column_map["final_settlement_amount"]
    ].map(parse_amount)

    if "quantity" in column_map:
        quantity = pd.to_numeric(raw_df[column_map["quantity"]], errors="coerce")
        missing_quantity_count = int(quantity.isna().sum())
        df["Quantity"] = quantity.fillna(1).clip(lower=0)
    else:
        missing_quantity_count = 0
        df["Quantity"] = 1.0

    for key, output_name in (
        ("sub_order_no", "Sub Order No"),
        ("order_date", "Order Date"),
        ("product_name", "Product Name"),
    ):
        if key in column_map:
            df[output_name] = raw_df[column_map[key]].values

    # The Meesho export includes a formula/notes row beneath the headings.
    # It has no recognized status and is safely excluded here.
    # Drop blank-status rows (including the Meesho formula/notes row).
    df = df[df["Live Order Status"].notna()].copy()
    df = df[df["Live Order Status"].astype(str).str.strip().ne("")].copy()
    df = df.reset_index(drop=True)
    df["_source_row"] = range(1, len(df) + 1)
    return df, column_map, [str(v) for v in excel_file.sheet_names]


def money(value: float) -> str:
    return f"₹{value:,.2f}"


def create_excel_report(
    summary_metrics: dict[str, float | int],
    sku_report: pd.DataFrame,
    delivered_detail: pd.DataFrame,
    return_detail: pd.DataFrame,
    rto_detail: pd.DataFrame,
    unknown_detail: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    summary_df = pd.DataFrame(
        [{"Metric": key, "Value": value} for key, value in summary_metrics.items()]
    )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Profit Summary", index=False)
        sku_report.to_excel(writer, sheet_name="SKU Profitability", index=False)
        delivered_detail.to_excel(writer, sheet_name="Delivered Orders", index=False)
        return_detail.to_excel(writer, sheet_name="Return Deductions", index=False)
        rto_detail.to_excel(writer, sheet_name="RTO Ignored", index=False)
        unknown_detail.to_excel(writer, sheet_name="Other Statuses", index=False)

        # Give each report sheet readable column widths and freeze its header.
        for worksheet in writer.book.worksheets:
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            for cells in worksheet.columns:
                max_length = max(
                    (len(str(cell.value)) for cell in cells if cell.value is not None),
                    default=10,
                )
                worksheet.column_dimensions[cells[0].column_letter].width = min(
                    max(max_length + 2, 12), 34
                )

    output.seek(0)
    return output.getvalue()


# ---------- Page UI ----------
st.title("Meesho Seller Profit Calculator")
st.caption(
    "Upload a Meesho payment workbook, enter your SKU purchase costs, "
    "and calculate profit after returns."
)

with st.expander("How this calculator works", expanded=False):
    st.markdown(
        """
        1. **Delivered:** profit = Final Settlement Amount − product cost.
        2. **Return:** the absolute value of each Return settlement amount is deducted from profit.
        3. **RTO:** ignored completely; neither settlement nor product cost is included.
        4. **Product cost:** enter the purchase cost per unit for each Supplier SKU. The app multiplies it by the order Quantity.

        The calculation uses the Final Settlement Amount from the payment sheet, not the listing price or sale amount.
        """
    )

uploaded_file = st.file_uploader(
    "Upload your Meesho payment Excel file",
    type=["xlsx", "xls", "xlsm"],
    help="The workbook must contain a sheet named 'Order Payments'.",
)

if uploaded_file is None:
    st.info("Upload your Meesho Excel file to begin.")
    st.stop()

file_bytes = uploaded_file.getvalue()
file_signature = hashlib.sha256(file_bytes).hexdigest()[:12]

try:
    orders, source_columns, sheet_names = load_order_payments(file_bytes)
except Exception as exc:
    st.error(str(exc))
    st.stop()

# Separate only the three statuses used by the calculator.
recognized = orders[orders["Live Order Status"].isin(VALID_STATUSES)].copy()
unknown_orders = orders[~orders["Live Order Status"].isin(VALID_STATUSES)].copy()

delivered = recognized[recognized["Live Order Status"] == "Delivered"].copy()
returns = recognized[recognized["Live Order Status"] == "Return"].copy()
rto = recognized[recognized["Live Order Status"] == "RTO"].copy()

# Remove rows without a SKU from delivered orders rather than calculating an
# apparently complete profit without a cost. Ask the user to correct the file.
missing_sku_rows = delivered[delivered["Supplier SKU"].eq("")]
if not missing_sku_rows.empty:
    st.error(
        f"{len(missing_sku_rows)} Delivered order(s) have a blank Supplier SKU. "
        "Please correct the SKU in the Excel file and upload it again."
    )
    st.dataframe(missing_sku_rows, use_container_width=True, hide_index=True)
    st.stop()

if delivered.empty and returns.empty and rto.empty:
    st.warning(
        "No Delivered, Return, or RTO orders were found. Check the Live Order Status values in your workbook."
    )
    if not unknown_orders.empty:
        st.write("Unrecognized statuses found:")
        st.dataframe(unknown_orders, use_container_width=True, hide_index=True)
    st.stop()

# Normalize numeric values and create SKU-level inputs from Delivered orders only.
delivered["Quantity"] = pd.to_numeric(delivered["Quantity"], errors="coerce").fillna(1).clip(lower=0)
delivered["Final Settlement Amount"] = pd.to_numeric(
    delivered["Final Settlement Amount"], errors="coerce"
).fillna(0.0)
returns["Final Settlement Amount"] = pd.to_numeric(
    returns["Final Settlement Amount"], errors="coerce"
).fillna(0.0)
rto["Final Settlement Amount"] = pd.to_numeric(
    rto["Final Settlement Amount"], errors="coerce"
).fillna(0.0)

sku_inputs = (
    delivered.groupby("Supplier SKU", dropna=False)
    .agg(
        **{
            "Delivered Orders": ("Supplier SKU", "size"),
            "Delivered Units": ("Quantity", "sum"),
        }
    )
    .reset_index()
    .sort_values("Supplier SKU", key=lambda s: s.astype(str).str.lower())
    .reset_index(drop=True)
)
sku_inputs["Unit Cost (₹)"] = None

st.subheader("1. Order summary")
metric_cols = st.columns(4)
metric_cols[0].metric("Delivered orders", f"{len(delivered):,}")
metric_cols[1].metric("Return orders", f"{len(returns):,}")
metric_cols[2].metric("RTO orders ignored", f"{len(rto):,}")
metric_cols[3].metric("Unique delivered SKUs", f"{len(sku_inputs):,}")

with st.expander("View detected statuses and source rows", expanded=False):
    status_summary = pd.DataFrame(
        [
            {
                "Status": "Delivered",
                "Orders": len(delivered),
                "Final Settlement Amount": delivered["Final Settlement Amount"].sum(),
                "Included in calculation": "Yes (minus SKU costs)",
            },
            {
                "Status": "Return",
                "Orders": len(returns),
                "Final Settlement Amount": returns["Final Settlement Amount"].sum(),
                "Included in calculation": "Yes (deducted)",
            },
            {
                "Status": "RTO",
                "Orders": len(rto),
                "Final Settlement Amount": rto["Final Settlement Amount"].sum(),
                "Included in calculation": "No (ignored)",
            },
        ]
    )
    st.dataframe(
        status_summary,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Final Settlement Amount": st.column_config.NumberColumn(format="₹%.2f")
        },
    )
    if not unknown_orders.empty:
        st.warning(
            f"{len(unknown_orders)} row(s) have blank or unrecognized Live Order Status values. "
            "They are not included in the calculation."
        )
        st.dataframe(unknown_orders, use_container_width=True, hide_index=True)

st.divider()
st.subheader("2. Enter the purchase cost for each Supplier SKU")
st.write(
    "Enter the **cost per unit** you pay for each SKU. The calculator multiplies this by "
    "the Delivered Quantity in the Excel file. Return and RTO SKUs do not need a cost entry."
)

if not delivered.empty:
    edited_sku_inputs = st.data_editor(
        sku_inputs,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key=f"sku_cost_editor_{file_signature}",
        column_config={
            "Supplier SKU": st.column_config.TextColumn(
                "Supplier SKU", help="SKU from the Order Payments sheet", disabled=True
            ),
            "Delivered Orders": st.column_config.NumberColumn(
                "Delivered Orders", disabled=True, format="%d"
            ),
            "Delivered Units": st.column_config.NumberColumn(
                "Delivered Units", disabled=True, format="%.2f"
            ),
            "Unit Cost (₹)": st.column_config.NumberColumn(
                "Unit Cost (₹)",
                help="Your purchase cost for one unit of this SKU, in rupees.",
                min_value=0.0,
                step=1.0,
                format="₹%.2f",
            ),
        },
    )
    entered_costs = pd.to_numeric(edited_sku_inputs["Unit Cost (₹)"], errors="coerce")
    missing_cost_mask = entered_costs.isna()
    missing_cost_skus = edited_sku_inputs.loc[missing_cost_mask, "Supplier SKU"].tolist()
else:
    edited_sku_inputs = sku_inputs
    entered_costs = pd.Series(dtype=float)
    missing_cost_skus = []

st.divider()
st.subheader("3. Profit calculation")

if missing_cost_skus:
    st.warning(
        f"Enter a unit cost for all {len(missing_cost_skus)} remaining SKU(s) to calculate final profit."
    )
    st.caption("Missing: " + ", ".join(str(sku) for sku in missing_cost_skus))
    st.stop()

# Map costs back to every Delivered order. Only Delivered orders incur product cost.
if not delivered.empty:
    cost_lookup = dict(
        zip(
            edited_sku_inputs["Supplier SKU"].astype(str),
            pd.to_numeric(edited_sku_inputs["Unit Cost (₹)"], errors="coerce").astype(float),
        )
    )
    delivered["Unit Cost (₹)"] = delivered["Supplier SKU"].map(cost_lookup).fillna(0.0)
    delivered["Product Cost (₹)"] = delivered["Unit Cost (₹)"] * delivered["Quantity"]
    delivered["Order Profit (₹)"] = (
        delivered["Final Settlement Amount"] - delivered["Product Cost (₹)"]
    )
else:
    delivered["Unit Cost (₹)"] = pd.Series(dtype=float)
    delivered["Product Cost (₹)"] = pd.Series(dtype=float)
    delivered["Order Profit (₹)"] = pd.Series(dtype=float)

# Returns are treated as a deduction by absolute amount. This prevents a negative
# settlement value (as seen in the uploaded sample) from increasing profit.
return_deduction = -float(returns["Final Settlement Amount"].abs().sum())
delivered_settlement = float(delivered["Final Settlement Amount"].sum())
delivered_product_cost = float(delivered["Product Cost (₹)"].sum()) if not delivered.empty else 0.0
delivered_profit = delivered_settlement - delivered_product_cost
net_profit = delivered_profit + return_deduction

if not delivered.empty:
    sku_report = (
        delivered.groupby("Supplier SKU", dropna=False)
        .agg(
            **{
                "Delivered Orders": ("Supplier SKU", "size"),
                "Delivered Units": ("Quantity", "sum"),
                "Final Settlement Amount (₹)": ("Final Settlement Amount", "sum"),
                "Product Cost (₹)": ("Product Cost (₹)", "sum"),
                "Profit Before Returns (₹)": ("Order Profit (₹)", "sum"),
            }
        )
        .reset_index()
    )
    sku_report["Unit Cost (₹)"] = sku_report["Supplier SKU"].map(cost_lookup)
    sku_report = sku_report[
        [
            "Supplier SKU",
            "Delivered Orders",
            "Delivered Units",
            "Unit Cost (₹)",
            "Final Settlement Amount (₹)",
            "Product Cost (₹)",
            "Profit Before Returns (₹)",
        ]
    ]
else:
    sku_report = pd.DataFrame(
        columns=[
            "Supplier SKU", "Delivered Orders", "Delivered Units", "Unit Cost (₹)",
            "Final Settlement Amount (₹)", "Product Cost (₹)", "Profit Before Returns (₹)"
        ]
    )

# Prepare order-level report tables.
base_columns = [
    name
    for name in ("Sub Order No", "Order Date", "Supplier SKU", "Product Name", "Quantity")
    if name in recognized.columns
]
delivered_detail = delivered[base_columns + [
    "Final Settlement Amount", "Unit Cost (₹)", "Product Cost (₹)", "Order Profit (₹)"
]].copy()
return_columns = base_columns + ["Final Settlement Amount"]
return_detail = returns[return_columns].copy() if not returns.empty else pd.DataFrame(columns=return_columns)
rto_detail = rto[return_columns].copy() if not rto.empty else pd.DataFrame(columns=return_columns)
unknown_detail = unknown_orders.copy()

summary_metrics: dict[str, float | int] = {
    "Delivered order count": int(len(delivered)),
    "Return order count": int(len(returns)),
    "RTO order count (ignored)": int(len(rto)),
    "Delivered final settlement total (₹)": delivered_settlement,
    "Delivered product cost total (₹)": delivered_product_cost,
    "Profit from Delivered orders before Returns (₹)": delivered_profit,
    "Return deduction (₹)": return_deduction,
    "NET PROFIT AFTER RETURNS (₹)": net_profit,
}

kpi_cols = st.columns(4)
kpi_cols[0].metric("Delivered settlement", money(delivered_settlement))
kpi_cols[1].metric("Delivered product cost", money(delivered_product_cost))
kpi_cols[2].metric("Return deduction", money(return_deduction))
kpi_cols[3].metric("Net profit after returns", money(net_profit))

st.caption(
    f"Calculation: {money(delivered_settlement)} delivered settlement − "
    f"{money(delivered_product_cost)} delivered product cost − "
    f"{money(abs(return_deduction))} return deduction = {money(net_profit)} net profit. "
    "RTO settlement amounts and RTO product costs are excluded."
)

st.markdown("#### SKU-wise profitability")
st.dataframe(
    sku_report,
    use_container_width=True,
    hide_index=True,
    column_config={
        col: st.column_config.NumberColumn(format="₹%.2f")
        for col in sku_report.columns
        if "(₹)" in col
    },
)

with st.expander("View Delivered order calculations", expanded=False):
    st.dataframe(
        delivered_detail,
        use_container_width=True,
        hide_index=True,
        column_config={
            col: st.column_config.NumberColumn(format="₹%.2f")
            for col in delivered_detail.columns
            if "(₹)" in col or col == "Final Settlement Amount" or col == "Order Profit (₹)"
        },
    )

with st.expander("View Return orders deducted", expanded=False):
    if return_detail.empty:
        st.write("No Return orders found.")
    else:
        st.dataframe(
            return_detail,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Final Settlement Amount": st.column_config.NumberColumn(format="₹%.2f")
            },
        )

with st.expander("View RTO orders ignored", expanded=False):
    if rto_detail.empty:
        st.write("No RTO orders found.")
    else:
        st.dataframe(
            rto_detail,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Final Settlement Amount": st.column_config.NumberColumn(format="₹%.2f")
            },
        )

report_bytes = create_excel_report(
    summary_metrics,
    sku_report,
    delivered_detail,
    return_detail,
    rto_detail,
    unknown_detail,
)
st.download_button(
    label="Download Profit Report (Excel)",
    data=report_bytes,
    file_name="meesho_profit_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

st.caption(
    "Privacy note: this app processes the uploaded workbook in the active session and does not save the original file to a database or permanent storage."
)
