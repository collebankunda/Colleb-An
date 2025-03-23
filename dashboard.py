import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
from io import BytesIO
from statsmodels.tsa.holtwinters import ExponentialSmoothing

st.title("Uganda Road Project Management Dashboard")

# Create navigation tabs
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Cost Analysis", "📉 CPI Graph", "📈 SPI",
    "⚠️ Risk Impact", "📅 Trend Analysis", "📂 Reports"
])

# Sidebar for file uploads
st.sidebar.header("Upload Data Files")
budget_file = st.sidebar.file_uploader("Upload Budgeted Cost File (Excel)", type=["xlsx"])
actual_file = st.sidebar.file_uploader("Upload Actual Cost File (Excel)", type=["xlsx"])
schedule_file = st.sidebar.file_uploader("Upload Schedule Data File (Excel)", type=["xlsx"])
risk_file = st.sidebar.file_uploader("Upload Risk Analysis File (Excel)", type=["xlsx"])

# Function to read Excel files
def read_excel(file):
    return pd.read_excel(file) if file else None

# Read data files
budget_df = read_excel(budget_file)
actual_df = read_excel(actual_file)
schedule_df = read_excel(schedule_file)
risk_df = read_excel(risk_file)

# Enhanced multiselect function with "Select All" support
def multiselect_with_select_all(label, options, default=None):
    full_options = ["Select All"] + options
    if default is None:
        default = ["Select All"]
    if "Select All" in default:
        default = full_options
    selected_options = st.sidebar.multiselect(label, full_options, default=default)
    if "Select All" in selected_options:
        return options
    return selected_options

if budget_df is not None and actual_df is not None:
    # Standardize column names
    budget_df.columns = budget_df.columns.str.strip()
    actual_df.columns = actual_df.columns.str.strip()

    if "Cost Category" in budget_df.columns and "Cost Category" in actual_df.columns:
        # Outer merge to include all categories from both files
        merged_df = pd.merge(
            budget_df,
            actual_df,
            on="Cost Category",
            suffixes=("_Budget", "_Actual"),
            how="outer"
        )
        # Fill missing values so calculations don't break
        merged_df["Amount_Budget"] = merged_df["Amount_Budget"].fillna(0)
        merged_df["Amount_Actual"] = merged_df["Amount_Actual"].fillna(0)
        # Calculate Variance and CPI safely
        merged_df["Variance"] = merged_df["Amount_Actual"] - merged_df["Amount_Budget"]

        def safe_cpi(row):
            if row["Amount_Actual"] == 0:
                return float("inf") if row["Amount_Budget"] > 0 else 1.0
            else:
                return row["Amount_Budget"] / row["Amount_Actual"]

        merged_df["CPI"] = merged_df.apply(safe_cpi, axis=1)

        # --- KPI Summary Section ---
        total_budget = merged_df["Amount_Budget"].sum()
        total_actual = merged_df["Amount_Actual"].sum()
        total_variance = merged_df["Variance"].sum()
        avg_cpi = merged_df["CPI"].mean()

        st.markdown("## Project Summary")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Budget (UGX)", f"{total_budget:,.0f}")
        col2.metric("Total Actual (UGX)", f"{total_actual:,.0f}")
        col3.metric("Total Variance (UGX)", f"{total_variance:,.0f}")
        col4.metric("Average CPI", f"{avg_cpi:.2f}")

        # --- Sidebar Interactive Filters ---
        st.sidebar.header("Interactive Filters")
        min_budget = float(merged_df["Amount_Budget"].min())
        max_budget = float(merged_df["Amount_Budget"].max())
        budget_threshold = st.sidebar.slider(
            "Select Budget Threshold (UGX)",
            min_value=int(min_budget),
            max_value=int(max_budget),
            value=int(min_budget)
        )

        cost_categories = merged_df["Cost Category"].unique().tolist()
        selected_categories = multiselect_with_select_all("Select Cost Categories", options=cost_categories)
        show_variance = st.sidebar.checkbox("Show Variance Analysis", value=True)

        # Filter data based on user input
        filtered_df = merged_df[
            (merged_df["Amount_Budget"] >= budget_threshold) &
            (merged_df["Cost Category"].isin(selected_categories))
        ]

        # --- TAB 1: Detailed Cost Comparison ---
        with tab1:
            st.subheader("Detailed Cost Comparison")
            st.dataframe(filtered_df)

        # --- TAB 2: Cost Performance Index (CPI) Graph ---
        with tab2:
            st.subheader("Cost Performance Index (CPI) Graph")
            if not filtered_df.empty:
                fig, ax = plt.subplots()
                ax.bar(filtered_df["Cost Category"], filtered_df["CPI"], color="blue")
                ax.axhline(1, color="red", linestyle="--", label="Baseline CPI = 1")
                ax.set_ylabel("CPI")
                ax.set_xticklabels(filtered_df["Cost Category"], rotation=45, ha="right")
                ax.legend()
                st.pyplot(fig)
            else:
                st.warning("No data available for the selected filters.")

        # --- TAB 3: Schedule Performance Index (SPI) ---
        with tab3:
            st.subheader("Schedule Performance Index (SPI)")
            if schedule_df is not None and {"Task", "Planned Duration", "Actual Duration"}.issubset(schedule_df.columns):
                schedule_df["SPI"] = schedule_df["Planned Duration"] / schedule_df["Actual Duration"]
                fig, ax = plt.subplots()
                ax.bar(schedule_df["Task"], schedule_df["SPI"], color="orange")
                ax.axhline(1, color="red", linestyle="--", label="Baseline SPI = 1")
                ax.set_ylabel("SPI")
                ax.set_xticklabels(schedule_df["Task"], rotation=45, ha="right")
                ax.legend()
                st.pyplot(fig)
            else:
                st.warning("Schedule data file must contain 'Task', 'Planned Duration', and 'Actual Duration' columns.")

        # --- TAB 4: Risk Impact Analysis ---
        with tab4:
            st.subheader("Risk Impact Analysis")
            if risk_df is not None and {"Risk Factor", "Variance"}.issubset(risk_df.columns):
                fig, ax = plt.subplots()
                ax.bar(risk_df["Risk Factor"], risk_df["Variance"], color="green")
                ax.set_ylabel("Variance (UGX)")
                ax.set_xticklabels(risk_df["Risk Factor"], rotation=45, ha="right")
                st.pyplot(fig)
            else:
                st.warning("Risk analysis file must contain 'Risk Factor' and 'Variance' columns.")

        # --- TAB 5: Quarterly Trend Analysis & Forecast ---
        with tab5:
            st.subheader("Quarterly Trend Analysis & Forecast")
            if "Date" in actual_df.columns:
                actual_df["Date"] = pd.to_datetime(actual_df["Date"])
                actual_df["Quarter"] = actual_df["Date"].dt.to_period("Q")
                quarterly_trend = actual_df.groupby("Quarter")["Amount_Actual"].sum().reset_index()
                quarterly_trend["Quarter"] = quarterly_trend["Quarter"].astype(str)
                fig, ax = plt.subplots()
                ax.plot(quarterly_trend["Quarter"], quarterly_trend["Amount_Actual"],
                        marker="o", linestyle="-", label="Actual")
                ts = quarterly_trend["Amount_Actual"]
                try:
                    model = ExponentialSmoothing(ts, trend="add", seasonal=None)
                    fit = model.fit()
                    forecast = fit.forecast(3)
                    forecast_quarters = [f"F{i+1}" for i in range(len(forecast))]
                    ax.plot(forecast_quarters, forecast, marker="o", linestyle="--", label="Forecast")
                except Exception as e:
                    st.error("Forecasting error: " + str(e))
                ax.set_xlabel("Quarter")
                ax.set_ylabel("Actual Cost (UGX)")
                ax.set_title("Quarterly Cost Trend Analysis & Forecast")
                ax.legend()
                st.pyplot(fig)
            else:
                st.warning("Actual cost data must contain a 'Date' column for trend analysis.")

        # --- TAB 6: Generate Reports ---
        with tab6:
            st.subheader("Generate Reports")
            def generate_pdf_report():
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=12)
                pdf.cell(200, 10, "Uganda Road Project Cost Analysis Report", ln=True, align="C")
                pdf.ln(10)
                for index, row in filtered_df.iterrows():
                    pdf.cell(
                        200,
                        10,
                        f"{row['Cost Category']}: UGX {row['Amount_Actual']:,} (Variance: UGX {row['Variance']:,})",
                        ln=True
                    )
                return pdf.output(dest="S").encode("latin1")
            st.download_button(label="📥 Download PDF Report",
                               data=generate_pdf_report(),
                               file_name="Cost_Analysis_Report.pdf",
                               mime="application/pdf")
            csv_data = merged_df.to_csv(index=False).encode("utf-8")
            st.download_button(label="📥 Download CSV Data",
                               data=csv_data,
                               file_name="Merged_Data.csv",
                               mime="text/csv")
    else:
        st.warning("The required column 'Cost Category' is missing in the uploaded files.")
else:
    st.warning("Please upload both Budgeted and Actual Cost files to proceed.")
