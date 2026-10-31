
from __future__ import annotations

import io
import json
import re
from pathlib import Path
from datetime import datetime, timezone

import pandas as pd
import streamlit as st

APP_TITLE = "Meesho Seller Profit Calculator"
COST_FILE = Path(__file__).parent / "sku_costs.json"
SUPABASE_TABLE = "sku_costs"

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="📊",
    layout="wide"
)


# -------------------- SETTINGS AND STORAGE --------------------

def get_secret(section: str, key: str):
    try:
        value = st.secrets.get(section, {}).get(key)
        return str(value).strip() if value else None
    except Exception:
        return None


def get_supabase_client():
    url = get_secret("supabase", "url")
    service_key = get_secret("supabase", "service_role_key")

    if not url or not service_key:
        return None

    from supabase import create_client
    return create_client(url, service_key)


def load_local_costs() -> dict[str, float]:
    if not COST_FILE.exists():
        return {}

    try:
        raw = json.loads(
            COST_FILE.read_text(encoding="utf-8")
        )

        return {
            str(sku): float(cost)
            for sku, cost in raw.items()
            if str(sku).strip() and float(cost) >= 0
        }

    except Exception:
        return {}


def load_costs():
    """Load previously saved SKU costs."""

    try:
        client = get_supabase_client()

        if client:
            result = (
                client.table(SUPABASE_TABLE)
                .select("sku, unit_cost")
                .order("sku")
                .execute()
            )

            costs = {
                str(row["sku"]): float(row["unit_cost"])
                for row in (result.data or [])
                if row.get("sku") is not None
                and row.get("unit_cost") is not None
            }

            return costs, "Supabase cloud database", None

        return load_local_costs(), "local JSON file", None

    except Exception as exc:
        return {}, "Supabase cloud database", str(exc)


def save_costs(costs_to_save: dict[str, float]):
    """Save new costs without deleting existing SKU costs."""

    clean = {
        str(sku).strip(): round(float(cost), 2)
        for sku, cost in costs_to_save.items()
        if str(sku).strip()
        and pd.notna(cost)
        and float(cost) >= 0
    }

    if not clean:
        return False, "No valid SKU costs were entered."

    try:
        client = get_supabase_client()

        if client:
            timestamp = datetime.now(timezone.utc).isoformat()

            rows = [
                {
                    "sku": sku,
                    "unit_cost": cost,
                    "updated_at": timestamp
                }
                for sku, cost in clean.items()
            ]

            client.table(SUPABASE_TABLE).upsert(
                rows,
                on_conflict="sku"
            ).execute()

            return True, (
                f"Saved {len(clean)} SKU cost(s) to Supabase."
            )

        # Local storage for running on your own computer.
        all_costs = load_local_costs()
        all_costs.update(clean)

        temp = COST_FILE.with_suffix(".tmp")

        temp.write_text(
            json.dumps(all_costs, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )

        temp.replace(COST_FILE)

        return True, (
            f"Saved {len(clean)} SKU cost(s) in {COST_FILE.name}."
        )

    except Exception as exc:
        return False, f"Could not save SKU costs: {exc}"


def require_password():
    """Optional app password."""

    expected = get_secret("app", "password")

    if not expected or st.session_state.get("logged_in", False):
        return

    st.title(APP_TITLE)
    st.info("Enter your app password to continue.")

    with st.form("login_form"):

        entered = st.text_input(
            "App password",
            type="password"
        )

        submitted = st.form_submit_button(
            "Unlock app",
            use_container_width=True
        )

    if submitted:

        if entered == expected:
            st.session_state["logged_in"] = True
            st.rerun()

        else:
            st.error("Incorrect password.")

    st.stop()


# -------------------- EXCEL READING --------------------

def normalize_header(value):
    return re.sub(
        r"[^a-z0-9]",
        "",
        str(value).strip().lower()
    )


def find_col(df, name):

    target = normalize_header(name)

    for col in df.columns:

        if normalize_header(col) == target:
            return col

    return None


def clean_sku(value):

    if pd.isna(value):
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value).strip()


def to_amount(value):

    if pd.isna(value):
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()

    if not text:
        return 0.0

    negative = text.startswith("(") and text.endswith(")")

    text = text.replace(",", "")
    text = text.replace("₹", "")
    text = text.replace("INR", "")
    text = text.replace("(", "").replace(")", "")

    text = re.sub(r"[^0-9.\-]", "", text)

    try:
        number = float(text)

        return -abs(number) if negative else number

    except ValueError:
        return 0.0


def get_status(value):

    if pd.isna(value):
        return None

    status = str(value).strip().lower()

    if not status:
        return None

    if "rto" in status or "return to origin" in status:
        return "RTO"

    if "return" in status:
        return "Return"

    if "deliver" in status:
        return "Delivered"

    return str(value).strip()


def read_meesho_excel(file_bytes):

    excel = pd.ExcelFile(io.BytesIO(file_bytes))

    sheet = next(
        (
            s for s in excel.sheet_names
            if s.strip().lower() == "order payments"
        ),
        None
    )

    if not sheet:
        raise ValueError(
            "Could not find the 'Order Payments' sheet. "
            "Sheets found: " + ", ".join(excel.sheet_names)
        )

    # Meesho has grouped headings on row 1
    # and actual column headings on row 2.

    raw = None

    for header_row in (1, 0, 2):

        candidate = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=sheet,
            header=header_row
        )

        required = (
            "Supplier SKU",
            "Live Order Status",
            "Final Settlement Amount"
        )

        if all(
            find_col(candidate, name) is not None
            for name in required
        ):

            raw = candidate.copy()
            break

    if raw is None:

        raise ValueError(
            "Could not find the required columns in Order Payments: "
            "Supplier SKU, Live Order Status, "
            "and Final Settlement Amount."
        )

    sku_col = find_col(raw, "Supplier SKU")
    status_col = find_col(raw, "Live Order Status")
    settlement_col = find_col(raw, "Final Settlement Amount")
    qty_col = find_col(raw, "Quantity")

    df = pd.DataFrame()

    df["Supplier SKU"] = raw[sku_col].map(clean_sku)

    df["Live Order Status"] = raw[status_col].map(get_status)

    df["Final Settlement Amount"] = (
        raw[settlement_col].map(to_amount)
    )

    if qty_col is not None:

        df["Quantity"] = (
            pd.to_numeric(raw[qty_col], errors="coerce")
            .fillna(1)
            .clip(lower=0)
        )

    else:

        df["Quantity"] = 1.0

    # Optional columns for the downloadable report.

    for name in (
        "Sub Order No",
        "Order Date",
        "Product Name"
    ):

        col = find_col(raw, name)

        if col is not None:
            df[name] = raw[col].values

    # Remove blank status rows and the Meesho formula/notes row.

    df = df[df["Live Order Status"].notna()].copy()

    df = df[
        df["Live Order Status"].astype(str).str.strip().ne("")
    ].reset_index(drop=True)

    return df


# -------------------- EXCEL REPORT --------------------

def make_excel_report(
    summary,
    sku_report,
    delivered,
    returns,
    rto,
    other
):

    output = io.BytesIO()

    summary_df = pd.DataFrame([
        {"Metric": k, "Value": v}
        for k, v in summary.items()
    ])

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        summary_df.to_excel(
            writer,
            sheet_name="Profit Summary",
            index=False
        )

        sku_report.to_excel(
            writer,
            sheet_name="SKU Profitability",
            index=False
        )

        delivered.to_excel(
            writer,
            sheet_name="Delivered Orders",
            index=False
        )

        returns.to_excel(
            writer,
            sheet_name="Return Deductions",
            index=False
        )

        rto.to_excel(
            writer,
            sheet_name="RTO Ignored",
            index=False
        )

        other.to_excel(
            writer,
            sheet_name="Other Statuses",
            index=False
        )

        for ws in writer.book.worksheets:

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for cells in ws.columns:

                width = max(
                    (
                        len(str(c.value))
                        for c in cells
                        if c.value is not None
                    ),
                    default=10
                )

                ws.column_dimensions[
                    cells[0].column_letter
                ].width = min(
                    max(width + 2, 12),
                    36
                )

    output.seek(0)

    return output.getvalue()


# -------------------- MAIN APPLICATION --------------------

require_password()

saved_costs, storage_mode, storage_error = load_costs()

st.title(APP_TITLE)

st.write(
    "Upload a Meesho payment Excel file, enter your purchase "
    "cost per Supplier SKU, and calculate your profit."
)

if storage_error:

    st.error(
        "Could not load SKU costs from Supabase. "
        "Check your Streamlit Secrets and Supabase table settings. "
        "Error: " + storage_error
    )

    st.stop()

elif storage_mode == "local JSON file":

    st.warning(
        "Local saving is active. This is suitable on your own "
        "computer. For permanent saving across devices and "
        "Streamlit Cloud restarts, configure Supabase below."
    )

else:

    st.success(
        "Saved SKU costs are loading from Supabase "
        "and will be shared across devices."
    )

if not get_secret("app", "password"):

    st.caption(
        "Security: set an app password in Streamlit Secrets "
        "before sharing a public app URL."
    )


# -------------------- SUPABASE INSTRUCTIONS --------------------

with st.expander(
    "One-time setup: permanent cloud saving with Supabase",
    expanded=False
):

    st.markdown("""
1. Create a free project at https://supabase.com/.
2. Open SQL Editor and run the SQL shown below.
3. In Supabase Project Settings → API, copy the Project URL
   and service_role key.
4. In Streamlit Community Cloud, open your app's
   Settings → Secrets and add the TOML shown below.
5. Save and restart the Streamlit app.

Never place the service role key in GitHub or in this Python file.
""")

    st.code("""
create table if not exists public.sku_costs (
    sku text primary key,
    unit_cost numeric(12, 2) not null
        check (unit_cost >= 0),
    updated_at timestamptz not null default now()
);

alter table public.sku_costs enable row level security;
""", language="sql")

    st.markdown("**Streamlit Secrets — replace the example values:**")

    st.code("""
[supabase]
url = "https://YOUR-PROJECT-ID.supabase.co"
service_role_key = "YOUR_SUPABASE_SERVICE_ROLE_KEY"

[app]
password = "choose-a-private-password"
""", language="toml")


# -------------------- SAVED SKU CATALOGUE --------------------

with st.expander(
    f"Saved SKU catalogue ({len(saved_costs)} SKUs)",
    expanded=False
):

    if saved_costs:

        catalogue = pd.DataFrame(
            sorted(saved_costs.items()),
            columns=[
                "Supplier SKU",
                "Saved Unit Cost (₹)"
            ]
        )

        st.dataframe(
            catalogue,
            use_container_width=True,
            hide_index=True
        )

    else:

        st.write(
            "No SKU costs saved yet. Upload your first report "
            "and enter the costs."
        )


# -------------------- UPLOAD EXCEL --------------------

uploaded = st.file_uploader(
    "Upload Meesho payment Excel file",
    type=["xlsx", "xls", "xlsm"]
)

if uploaded is None:

    st.info("Upload an Excel file to start calculating.")

    st.stop()

try:

    orders = read_meesho_excel(
        uploaded.getvalue()
    )

except Exception as exc:

    st.error(str(exc))

    st.stop()


# -------------------- CLASSIFY ORDERS --------------------

delivered = orders[
    orders["Live Order Status"] == "Delivered"
].copy()

returns = orders[
    orders["Live Order Status"] == "Return"
].copy()

rto = orders[
    orders["Live Order Status"] == "RTO"
].copy()

other = orders[
    ~orders["Live Order Status"].isin([
        "Delivered",
        "Return",
        "RTO"
    ])
].copy()


if (delivered["Supplier SKU"] == "").any():

    st.error(
        "One or more Delivered orders have a blank Supplier SKU. "
        "Correct the Excel file and upload it again."
    )

    st.dataframe(
        delivered[delivered["Supplier SKU"] == ""],
        use_container_width=True
    )

    st.stop()


# -------------------- ORDER SUMMARY --------------------

st.subheader("1. Order summary")

c1, c2, c3, c4 = st.columns(4)

c1.metric(
    "Delivered orders",
    f"{len(delivered):,}"
)

c2.metric(
    "Return orders",
    f"{len(returns):,}"
)

c3.metric(
    "RTO orders ignored",
    f"{len(rto):,}"
)

c4.metric(
    "Delivered SKUs",
    f"{delivered['Supplier SKU'].nunique():,}"
)

if not other.empty:

    with st.expander(
        f"Other or unrecognized statuses ({len(other)} orders)"
    ):

        st.dataframe(
            other,
            use_container_width=True,
            hide_index=True
        )


# -------------------- SKU COST ENTRY --------------------

st.subheader("2. Enter or confirm Supplier SKU costs")

st.caption(
    "Enter the cost for ONE unit of each SKU. "
    "Previously saved costs are filled in automatically. "
    "Delivered Quantity is multiplied by the unit cost. "
    "Returns and RTO do not need a product cost entry."
)

if not delivered.empty:

    sku_table = (
        delivered.groupby(
            "Supplier SKU",
            as_index=False
        )
        .agg(
            Delivered_Orders=("Supplier SKU", "size"),
            Delivered_Units=("Quantity", "sum")
        )
        .rename(columns={
            "Delivered_Orders": "Delivered Orders",
            "Delivered_Units": "Delivered Units"
        })
    )

    sku_table["Unit Cost (₹)"] = (
        sku_table["Supplier SKU"].map(saved_costs)
    )

    sku_table = (
        sku_table.sort_values("Supplier SKU")
        .reset_index(drop=True)
    )

    edited = st.data_editor(
        sku_table,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key=(
            "sku_cost_editor_"
            + str(abs(hash(uploaded.name + str(uploaded.size))))
        ),
        column_config={

            "Supplier SKU": st.column_config.TextColumn(
                disabled=True
            ),

            "Delivered Orders": st.column_config.NumberColumn(
                disabled=True
            ),

            "Delivered Units": st.column_config.NumberColumn(
                disabled=True,
                format="%.2f"
            ),

            "Unit Cost (₹)": st.column_config.NumberColumn(
                min_value=0.0,
                step=1.0,
                format="₹%.2f",
                help="Your purchase cost for one unit of this SKU."
            )
        }
    )

    entered = pd.to_numeric(
        edited["Unit Cost (₹)"],
        errors="coerce"
    )

    save_col, info_col = st.columns([1, 3])

    with save_col:

        save_clicked = st.button(
            "Save SKU Costs",
            type="primary",
            use_container_width=True
        )

    with info_col:

        st.caption(
            "Click Save SKU Costs after entering or changing costs. "
            "Costs are reused for later Excel uploads."
        )

    if save_clicked:

        to_save = {
            str(sku): float(cost)
            for sku, cost in zip(
                edited["Supplier SKU"],
                entered
            )
            if pd.notna(cost)
        }

        ok, message = save_costs(to_save)

        if ok:

            saved_costs.update(to_save)

            st.success(message)

        else:

            st.error(message)

    missing_skus = edited.loc[
        entered.isna(),
        "Supplier SKU"
    ].astype(str).tolist()

else:

    edited = pd.DataFrame(
        columns=["Supplier SKU", "Unit Cost (₹)"]
    )

    entered = pd.Series(dtype=float)

    missing_skus = []

    st.info(
        "No Delivered orders in this workbook, "
        "so no SKU cost entry is required."
    )


# -------------------- PROFIT CALCULATION --------------------

st.subheader("3. Profit calculation")

if missing_skus:

    st.warning(
        f"Enter a cost for all {len(missing_skus)} missing SKU(s) "
        "to calculate final profit. You can save the costs you "
        "have entered and finish the remaining SKUs later."
    )

    st.write(
        "Missing costs: " + ", ".join(missing_skus)
    )

    st.stop()


cost_map = {
    str(sku): float(cost)
    for sku, cost in zip(
        edited.get("Supplier SKU", []),
        entered
    )
    if pd.notna(cost)
}


# Delivered orders: settlement minus product cost.

if not delivered.empty:

    delivered["Unit Cost (₹)"] = (
        delivered["Supplier SKU"]
        .map(cost_map)
        .fillna(0.0)
    )

    delivered["Product Cost (₹)"] = (
        delivered["Unit Cost (₹)"]
        * delivered["Quantity"]
    )

    delivered["Order Profit (₹)"] = (
        delivered["Final Settlement Amount"]
        - delivered["Product Cost (₹)"]
    )

else:

    for col in (
        "Unit Cost (₹)",
        "Product Cost (₹)",
        "Order Profit (₹)"
    ):

        delivered[col] = pd.Series(dtype=float)


# Returns: deduct the absolute settlement amount.
# RTO: completely excluded.

return_deduction = float(
    returns["Final Settlement Amount"].abs().sum()
)

delivered_settlement = float(
    delivered["Final Settlement Amount"].sum()
)

delivered_cost = (
    float(delivered["Product Cost (₹)"].sum())
    if not delivered.empty
    else 0.0
)

delivered_profit = (
    delivered_settlement - delivered_cost
)

net_profit = (
    delivered_profit - return_deduction
)


# -------------------- SKU-WISE REPORT --------------------

if not delivered.empty:

    sku_report = delivered.groupby(
        "Supplier SKU",
        as_index=False
    ).agg(
        **{
            "Delivered Orders": ("Supplier SKU", "size"),
            "Delivered Units": ("Quantity", "sum"),
            "Final Settlement Amount (₹)": (
                "Final Settlement Amount", "sum"
            ),
            "Product Cost (₹)": (
                "Product Cost (₹)", "sum"
            ),
            "Profit Before Returns (₹)": (
                "Order Profit (₹)", "sum"
            )
        }
    )

    sku_report["Unit Cost (₹)"] = (
        sku_report["Supplier SKU"].map(cost_map)
    )

    sku_report = sku_report[
        [
            "Supplier SKU",
            "Delivered Orders",
            "Delivered Units",
            "Unit Cost (₹)",
            "Final Settlement Amount (₹)",
            "Product Cost (₹)",
            "Profit Before Returns (₹)"
        ]
    ]

else:

    sku_report = pd.DataFrame(
        columns=[
            "Supplier SKU",
            "Delivered Orders",
            "Delivered Units",
            "Unit Cost (₹)",
            "Final Settlement Amount (₹)",
            "Product Cost (₹)",
            "Profit Before Returns (₹)"
        ]
    )


# -------------------- ORDER LEVEL DETAILS --------------------

base_cols = [
    c for c in [
        "Sub Order No",
        "Order Date",
        "Supplier SKU",
        "Product Name",
        "Quantity"
    ]
    if c in orders.columns
]

delivered_report = delivered[
    base_cols + [
        "Final Settlement Amount",
        "Unit Cost (₹)",
        "Product Cost (₹)",
        "Order Profit (₹)"
    ]
].copy()

return_report = returns[
    base_cols + ["Final Settlement Amount"]
].copy()

rto_report = rto[
    base_cols + ["Final Settlement Amount"]
].copy()


# -------------------- SUMMARY --------------------

summary = {

    "Delivered order count": len(delivered),

    "Return order count": len(returns),

    "RTO order count (ignored)": len(rto),

    "Delivered settlement total (₹)": delivered_settlement,

    "Delivered product cost total (₹)": delivered_cost,

    "Profit before Returns (₹)": delivered_profit,

    "Return deduction (₹)": return_deduction,

    "NET PROFIT AFTER RETURNS (₹)": net_profit
}


# -------------------- FINAL PROFIT DASHBOARD --------------------

st.divider()

st.subheader("4. Your profit")

k1, k2, k3, k4 = st.columns(4)

k1.metric(
    "Delivered settlement",
    f"₹{delivered_settlement:,.2f}"
)

k2.metric(
    "Product cost",
    f"₹{delivered_cost:,.2f}"
)

k3.metric(
    "Return deduction",
    f"-₹{return_deduction:,.2f}"
)

k4.metric(
    "Net profit after returns",
    f"₹{net_profit:,.2f}"
)

st.caption(
    f"₹{delivered_settlement:,.2f} delivered settlement "
    f"− ₹{delivered_cost:,.2f} product cost "
    f"− ₹{return_deduction:,.2f} Return deduction "
    f"= ₹{net_profit:,.2f}. "
    "RTO settlement amounts are ignored completely."
)


# -------------------- DISPLAY REPORTS --------------------

st.markdown("#### SKU-wise profitability")

st.dataframe(
    sku_report,
    use_container_width=True,
    hide_index=True
)

with st.expander("View Delivered order calculations"):

    st.dataframe(
        delivered_report,
        use_container_width=True,
        hide_index=True
    )

with st.expander("View Return orders deducted"):

    st.dataframe(
        return_report,
        use_container_width=True,
        hide_index=True
    )

with st.expander("View RTO orders ignored"):

    st.dataframe(
        rto_report,
        use_container_width=True,
        hide_index=True
    )


# -------------------- DOWNLOAD EXCEL REPORT --------------------

report = make_excel_report(
    summary,
    sku_report,
    delivered_report,
    return_report,
    rto_report,
    other
)

st.download_button(
    "Download Profit Report (Excel)",
    data=report,
    file_name="meesho_profit_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)

st.caption(
    "Your uploaded Excel file is processed in memory and is not stored "
    "by this app. Only Supplier SKU names and their saved unit costs "
    "are kept in the cost database."
)
