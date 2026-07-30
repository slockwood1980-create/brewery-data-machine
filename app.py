"""
Simon's Magic Data Machine
---------------------------
A Streamlit dashboard that turns a Delphi (Salesforce) events-booking
export into financial and pipeline KPIs for The Brewery.

Run locally:   streamlit run app.py
Deploy:        see README.md
"""

import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# --------------------------------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Simon's Magic Data Machine",
    page_icon="🍺",
    layout="wide",
)

# --------------------------------------------------------------------------
# CONSTANTS — column mapping & business rules
# --------------------------------------------------------------------------
# Maps the raw Delphi/Salesforce export column names to clean internal names.
# If your export ever renames a column slightly, add the new name here.
COLUMN_MAP = {
    "Booking: Booking Post As": "Booking_ID",
    "Account": "Account",
    "Status": "Status",
    "Blended Revenue Total Currency": "Currency",
    "Blended Revenue Total": "Revenue",
    "Arrival": "Arrival_Date",
    "Booking: Created Date": "Created_Date",
    "Booking: Owner Alias": "Sales_Rep",
    "Date Definite": "Date_Definite",
    "Date Lost": "Date_Lost",
    "Date Tentative": "Date_Tentative",
    "Lead Source": "Lead_Source",
    "Meeting Class": "Meeting_Class",
    "Date Turned Down": "Date_Turned_Down",
    "Lost Reason": "Lost_Reason",
}

DATE_COLS = [
    "Arrival_Date", "Created_Date", "Date_Definite",
    "Date_Lost", "Date_Tentative", "Date_Turned_Down",
]

# Business rules confirmed with Simon:
#   - Won            = Definite
#   - Lost (excluded)= Lost, TurnedDown, Cancelled  (Cancelled counts as lost revenue)
#   - Open pipeline  = Tentative, Prospect (Prospect == "Provisional")
WON_STATUS = "Definite"
LOST_STATUSES = ["Lost", "TurnedDown", "Cancelled"]
OPEN_STATUSES = ["Tentative", "Prospect"]


# --------------------------------------------------------------------------
# DATA LOADING
# --------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_data(uploaded_file) -> pd.DataFrame:
    """
    Reads the uploaded file. Delphi/Salesforce 'xls' exports are frequently
    actually HTML tables wearing an .xls extension, so we try real Excel
    formats first, then fall back to HTML parsing.
    """
    # Try modern Excel (.xlsx)
    try:
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        return df
    except Exception:
        pass

    # Try legacy Excel (.xls, true binary format)
    try:
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, engine="xlrd")
        return df
    except Exception:
        pass

    # Fall back to HTML-disguised-as-xls (common Salesforce/Delphi export)
    try:
        uploaded_file.seek(0)
        tables = pd.read_html(uploaded_file)
        if tables:
            return tables[0]
    except Exception:
        pass

    raise ValueError(
        "Could not read this file. Please make sure it's a Delphi export "
        "in .xlsx or .xls format."
    )


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Renames columns, parses dates, and derives helper fields."""
    # Rename any columns we recognise; leave others untouched
    rename_dict = {k: v for k, v in COLUMN_MAP.items() if k in df.columns}
    df = df.rename(columns=rename_dict)

    missing = [v for v in COLUMN_MAP.values() if v not in df.columns]
    if missing:
        st.warning(
            f"Heads up — these expected columns weren't found in your file "
            f"and related metrics will be skipped: {', '.join(missing)}"
        )

    # Parse dates (Delphi exports as dd/mm/yyyy)
    for col in DATE_COLS:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], format="%d/%m/%Y", errors="coerce")

    # Ensure Revenue is numeric
    if "Revenue" in df.columns:
        df["Revenue"] = pd.to_numeric(df["Revenue"], errors="coerce").fillna(0)

    # Derive Financial Year from Arrival Date (FY runs 1 July - 30 June)
    if "Arrival_Date" in df.columns:
        df["Financial_Year"] = df["Arrival_Date"].apply(fy_from_date)

    # Derive a simple outcome bucket for easy filtering/charting
    if "Status" in df.columns:
        df["Outcome"] = df["Status"].apply(bucket_status)

    return df


def fy_from_date(d) -> str:
    """Converts a date into a 'FY2025/26' style label (FY = 1 Jul - 30 Jun)."""
    if pd.isna(d):
        return "Unknown"
    start_year = d.year if d.month >= 7 else d.year - 1
    return f"FY{start_year}/{str(start_year + 1)[-2:]}"


def bucket_status(status: str) -> str:
    if status == WON_STATUS:
        return "Won"
    if status in LOST_STATUSES:
        return "Lost"
    if status in OPEN_STATUSES:
        return "Open"
    return "Other"


# --------------------------------------------------------------------------
# KPI CALCULATIONS
# --------------------------------------------------------------------------
def calculate_kpis(df: pd.DataFrame) -> dict:
    k = {}

    total_leads = len(df)
    won_df = df[df["Status"] == WON_STATUS]
    lost_df = df[df["Status"].isin(LOST_STATUSES)]
    open_df = df[df["Status"].isin(OPEN_STATUSES)]

    won_revenue = won_df["Revenue"].sum()
    total_pipeline_value = df["Revenue"].sum()

    # --- Core metrics ---
    k["total_leads"] = total_leads
    k["won_revenue"] = won_revenue
    k["total_pipeline_value"] = total_pipeline_value
    k["revenue_per_lead"] = won_revenue / total_leads if total_leads else 0
    k["pipeline_conversion"] = (
        won_revenue / total_pipeline_value if total_pipeline_value else 0
    )
    k["conversion_rate"] = len(won_df) / total_leads if total_leads else 0

    # --- Additional KPIs ---
    # 1. Average Deal Size (Won only)
    k["avg_deal_size"] = won_revenue / len(won_df) if len(won_df) else 0

    # 2. Sales Cycle Length: enquiry -> signed contract, for Won deals
    cycle_df = won_df.dropna(subset=["Created_Date", "Date_Definite"])
    if len(cycle_df):
        cycle_days = (cycle_df["Date_Definite"] - cycle_df["Created_Date"]).dt.days
        k["avg_sales_cycle_days"] = cycle_days.mean()
    else:
        k["avg_sales_cycle_days"] = None

    # 3. Lost vs Turned Down split (both are "not won" but mean different things)
    lost_count = (df["Status"] == "Lost").sum()
    turned_down_count = (df["Status"] == "TurnedDown").sum()
    not_won_total = lost_count + turned_down_count
    k["lost_rate"] = lost_count / not_won_total if not_won_total else 0
    k["turned_down_rate"] = turned_down_count / not_won_total if not_won_total else 0
    k["lost_count"] = lost_count
    k["turned_down_count"] = turned_down_count

    # 4. Cancellation Rate: signed deals that later fell through
    cancelled_count = (df["Status"] == "Cancelled").sum()
    signed_total = len(won_df) + cancelled_count
    k["cancellation_rate"] = cancelled_count / signed_total if signed_total else 0
    k["cancelled_count"] = cancelled_count

    # 5. Win rate by Lead Source
    if "Lead_Source" in df.columns:
        by_source = df.groupby("Lead_Source").agg(
            Total_Leads=("Status", "count"),
            Won=("Status", lambda s: (s == WON_STATUS).sum()),
            Revenue_Won=("Revenue", lambda s: s[df.loc[s.index, "Status"] == WON_STATUS].sum()),
        )
        by_source["Win_Rate"] = by_source["Won"] / by_source["Total_Leads"]
        k["by_lead_source"] = by_source.sort_values("Revenue_Won", ascending=False)

    # 6. Revenue by Meeting Class (event type) — Won only
    if "Meeting_Class" in df.columns:
        by_class = won_df.groupby("Meeting_Class").agg(
            Won_Deals=("Status", "count"),
            Revenue_Won=("Revenue", "sum"),
        ).sort_values("Revenue_Won", ascending=False)
        k["by_meeting_class"] = by_class

    # 7. Pipeline Aging — open bookings bucketed by days since enquiry
    if len(open_df):
        today = pd.Timestamp(datetime.now().date())
        aging = (today - open_df["Created_Date"]).dt.days
        bins = [-1, 30, 60, 90, 10_000]
        labels = ["0-30 days", "31-60 days", "61-90 days", "90+ days"]
        aging_bucket = pd.cut(aging, bins=bins, labels=labels)
        k["pipeline_aging"] = (
            open_df.assign(Aging_Bucket=aging_bucket)
            .groupby("Aging_Bucket", observed=True)
            .agg(Open_Deals=("Status", "count"), Open_Value=("Revenue", "sum"))
        )
    else:
        k["pipeline_aging"] = None

    return k


# --------------------------------------------------------------------------
# UI HELPERS
# --------------------------------------------------------------------------
def gbp(value) -> str:
    if value is None or pd.isna(value):
        return "—"
    return f"£{value:,.0f}"


def pct(value) -> str:
    if value is None or pd.isna(value):
        return "—"
    return f"{value * 100:.1f}%"


# --------------------------------------------------------------------------
# MAIN APP
# --------------------------------------------------------------------------
def main():
    st.title("🍺 Simon's Magic Data Machine")
    st.caption(
        "Upload your Delphi booking export to see The Brewery's sales "
        "pipeline and revenue KPIs — instantly."
    )

    uploaded_file = st.file_uploader(
        "Upload your Delphi export (.xlsx or .xls)", type=["xlsx", "xls"]
    )

    if uploaded_file is None:
        st.info("👆 Upload a file to get started.")
        st.stop()

    with st.spinner("Reading your file..."):
        try:
            raw_df = load_data(uploaded_file)
            df = clean_data(raw_df)
        except Exception as e:
            st.error(f"Something went wrong reading this file: {e}")
            st.stop()

    # ---------------- Sidebar filters ----------------
    st.sidebar.header("Filters")

    fy_options = ["All Time"] + sorted(
        [fy for fy in df["Financial_Year"].dropna().unique() if fy != "Unknown"],
        reverse=True,
    )
    selected_fy = st.sidebar.selectbox("Financial Year (based on Arrival date)", fy_options)

    if selected_fy != "All Time":
        df = df[df["Financial_Year"] == selected_fy]

    if "Sales_Rep" in df.columns:
        reps = ["All"] + sorted(df["Sales_Rep"].dropna().unique().tolist())
        selected_rep = st.sidebar.selectbox("Sales Person", reps)
        if selected_rep != "All":
            df = df[df["Sales_Rep"] == selected_rep]

    if len(df) == 0:
        st.warning("No bookings match the selected filters.")
        st.stop()

    kpis = calculate_kpis(df)

    # ---------------- Core Metrics ----------------
    st.header("Core Metrics")
    c1, c2, c3 = st.columns(3)
    c1.metric(
        "Revenue per Lead",
        gbp(kpis["revenue_per_lead"]),
        help="Won Revenue ÷ Total Number of Bookings (leads) in the selected period.",
    )
    c2.metric(
        "Pipeline Value Conversion",
        pct(kpis["pipeline_conversion"]),
        help="Actual Revenue Won ÷ Total Pipeline Value (sum of all booking values, Won + Lost + Open).",
    )
    c3.metric(
        "Conversion Rate",
        pct(kpis["conversion_rate"]),
        help="Number of Definite (Won) bookings ÷ Total Number of Bookings (leads).",
    )

    st.caption(
        f"Based on {kpis['total_leads']:,} bookings totalling "
        f"{gbp(kpis['total_pipeline_value'])} in pipeline value, of which "
        f"{gbp(kpis['won_revenue'])} was won."
    )

    st.divider()

    # ---------------- Additional KPIs ----------------
    st.header("Additional KPIs")
    c1, c2, c3 = st.columns(3)
    c1.metric("Average Deal Size (Won)", gbp(kpis["avg_deal_size"]))
    c2.metric(
        "Avg. Sales Cycle",
        f"{kpis['avg_sales_cycle_days']:.0f} days" if kpis["avg_sales_cycle_days"] is not None else "—",
        help="Average days from enquiry (Booking Created Date) to signed contract (Date Definite).",
    )
    c3.metric(
        "Cancellation Rate",
        pct(kpis["cancellation_rate"]),
        help="Of all signed deals (Definite + Cancelled), the % that later cancelled.",
    )

    c4, c5 = st.columns(2)
    c4.metric(
        "Lost to Competitor Rate",
        pct(kpis["lost_rate"]),
        f"{kpis['lost_count']:,} bookings",
        help="Of all non-won outcomes (Lost + TurnedDown), the % where the client chose someone else.",
    )
    c5.metric(
        "Turned Down Rate",
        pct(kpis["turned_down_rate"]),
        f"{kpis['turned_down_count']:,} bookings",
        help="Of all non-won outcomes (Lost + TurnedDown), the % where The Brewery declined to host.",
    )

    st.divider()

    # ---------------- Charts ----------------
    st.header("Breakdowns")

    tab1, tab2, tab3 = st.tabs(["By Lead Source", "By Event Type", "Pipeline Aging"])

    with tab1:
        if "by_lead_source" in kpis:
            source_df = kpis["by_lead_source"].reset_index()
            fig = px.bar(
                source_df,
                x="Lead_Source",
                y="Revenue_Won",
                hover_data=["Total_Leads", "Won", "Win_Rate"],
                title="Revenue Won by Lead Source",
                labels={"Revenue_Won": "Revenue Won (£)", "Lead_Source": "Lead Source"},
            )
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(
                source_df.style.format({
                    "Revenue_Won": "£{:,.0f}",
                    "Win_Rate": "{:.1%}",
                }),
                use_container_width=True,
            )
        else:
            st.info("Lead Source column not found in this file.")

    with tab2:
        if "by_meeting_class" in kpis:
            class_df = kpis["by_meeting_class"].reset_index()
            fig = px.bar(
                class_df,
                x="Meeting_Class",
                y="Revenue_Won",
                hover_data=["Won_Deals"],
                title="Revenue Won by Event Type",
                labels={"Revenue_Won": "Revenue Won (£)", "Meeting_Class": "Event Type"},
            )
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(
                class_df.style.format({"Revenue_Won": "£{:,.0f}"}),
                use_container_width=True,
            )
        else:
            st.info("Meeting Class column not found in this file.")

    with tab3:
        if kpis["pipeline_aging"] is not None:
            aging_df = kpis["pipeline_aging"].reset_index()
            fig = px.bar(
                aging_df,
                x="Aging_Bucket",
                y="Open_Value",
                hover_data=["Open_Deals"],
                title="Open Pipeline Value by Age (Tentative + Prospect bookings)",
                labels={"Open_Value": "Open Pipeline Value (£)", "Aging_Bucket": "Days Since Enquiry"},
            )
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(
                aging_df.style.format({"Open_Value": "£{:,.0f}"}),
                use_container_width=True,
            )
        else:
            st.info("No open (Tentative/Prospect) bookings in this selection.")

    st.divider()

    # ---------------- Raw data preview ----------------
    with st.expander("View raw data"):
        st.dataframe(df, use_container_width=True)


if __name__ == "__main__":
    main()
