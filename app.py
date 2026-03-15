import streamlit as st
import pandas as pd
from datetime import date
from datetime import timedelta

st.set_page_config(page_title="Depreciation Calculator", layout="wide")

# ----------------------------------------------------
# EXPECTED FAR FORMAT
# ----------------------------------------------------
EXPECTED_COLUMNS = [
    "Asset ID",
    "Asset Name",
    "Asset Category",
    "Cost",
    "Rate",
    "Acquisition Date"
]


# ----------------------------------------------------
# STRAIGHT LINE CALCULATOR
# ----------------------------------------------------
def calculator(cost, acquisition_date, period_start, period_end, rate):

    ul_days = int((1 / rate) * 365)
    eol = acquisition_date + timedelta(days=ul_days)
    dep_daily = (rate * cost) / 365

    if period_start >= eol:
        return 0.0

    dep_start = max(acquisition_date, period_start)
    dep_end = min(period_end, eol)

    dep_days = (dep_end - dep_start).days + 1

    if dep_days <= 0:
        return 0.0

    return dep_daily * dep_days


# ----------------------------------------------------
# REDUCING BALANCE CALCULATOR
# ----------------------------------------------------
def calculator2(cost, acquisition_date, period_start, period_end, rate):

    ul_days = int((1 / rate) * 365)
    eol = acquisition_date + timedelta(days=ul_days)

    if period_start >= eol:
        return 0.0

    dep_start = max(acquisition_date, period_start)
    dep_end = min(period_end, eol)

    dep_days = (dep_end - dep_start).days + 1

    if dep_days <= 0:
        return 0.0

    days_before_period = (dep_start - acquisition_date).days
    years_elapsed = days_before_period / 365

    book_value = cost * ((1 - rate) ** years_elapsed)

    dep_daily = (rate * book_value) / 365

    return dep_daily * dep_days


# ----------------------------------------------------
# DEPRECIATION ENGINE
# ----------------------------------------------------
def calculate_depreciation(far_df, period_start, period_end, method):

    results_list = []
    category_totals = {}

    for _, row in far_df.iterrows():

        asset_id = row["Asset ID"]
        asset_name = row["Asset Name"]
        category = row["Asset Category"]
        cost = row["Cost"]
        rate = row["Rate"]

        acquisition_date = row["Acquisition Date"].date()

        if method == "Straight Line":
            depreciation = calculator(
                cost,
                acquisition_date,
                period_start,
                period_end,
                rate
            )
            if depreciation is None:
                depreciation = 0.0
        else:
            depreciation = calculator2(
                cost,
                acquisition_date,
                period_start,
                period_end,
                rate
            )
            if depreciation is None:
                depreciation = 0.0

        results_list.append({
            "Asset ID": asset_id,
            "Asset Name": asset_name,
            "Asset Category": category,
            "Cost": cost,
            "Depreciation Charge": depreciation
        })

        if category not in category_totals:
            category_totals[category] = 0

        category_totals[category] += depreciation

    results = pd.DataFrame(results_list)

    summary = pd.DataFrame(
        category_totals.items(),
        columns=["Asset Category", "Total Depreciation"]
    )

    return results, summary


# ----------------------------------------------------
# UI
# ----------------------------------------------------
st.title("Asset Depreciation Calculator")

st.write(
    "Upload the Fixed Asset Register (FAR) in the format of the template to calculate depreciation."
)


# template download
with open("template.xlsx", "rb") as file:
    template_bytes = file.read()

st.download_button(
    label="Download FAR Template",
    data=template_bytes,
    file_name="FAR_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


# upload FAR
uploaded_file = st.file_uploader(
    "Upload FAR file in the format of the template",
    type=["xlsx"]
)


if uploaded_file:

    far_df = pd.read_excel(uploaded_file)
    far_df = far_df.dropna(how="all")

    # ----------------------------------------------------
    # VALIDATE COLUMN FORMAT
    # ----------------------------------------------------
    uploaded_columns = list(far_df.columns)

    missing_columns = set(EXPECTED_COLUMNS) - set(uploaded_columns)
    extra_columns = set(uploaded_columns) - set(EXPECTED_COLUMNS)

    if missing_columns or extra_columns:

        st.error(
            "The uploaded FAR does not match the required template format. "
            "Please download and use the provided template."
        )

        if missing_columns:
            st.write("Missing columns:", list(missing_columns))

        if extra_columns:
            st.write("Unexpected columns:", list(extra_columns))

        st.stop()

    # ----------------------------------------------------
    # DATE PARSING
    # ----------------------------------------------------
    far_df["Acquisition Date"] = pd.to_datetime(
        far_df["Acquisition Date"],
        dayfirst=True
    )

    st.subheader("Uploaded FAR")
    st.dataframe(far_df, use_container_width=True)

    # depreciation period
    st.subheader("Select Depreciation Period")

    col1, col2 = st.columns(2)

    with col1:
        period_start = st.date_input("Period Start")

    with col2:
        period_end = st.date_input("Period End")

    # method dropdown
    method = st.selectbox(
        "Depreciation Method",
        ["Straight Line", "Reducing Balance"]
    )

    # run depreciation
    if st.button("Run Depreciation"):

        results, summary = calculate_depreciation(
            far_df,
            period_start,
            period_end,
            method
        )

        st.success("Depreciation calculation complete.")

        st.subheader("Depreciation Summary")
        st.dataframe(summary, use_container_width=True)

        summary_csv = summary.to_csv(index=False)
        results_csv = results.to_csv(index=False)

        col1, col2 = st.columns(2)

        with col1:
            st.download_button(
                label="Download Summary",
                data=summary_csv,
                file_name="depreciation_summary.csv",
                mime="text/csv"
            )

        with col2:
            st.download_button(
                label="Download Full FAR with Depreciation",
                data=results_csv,
                file_name="far_with_depreciation.csv",
                mime="text/csv"
            )
