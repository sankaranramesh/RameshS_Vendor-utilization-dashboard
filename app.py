
import streamlit as st
import pandas as pd
import numpy as np
import calendar
from io import BytesIO

st.set_page_config(page_title="Vendor Forecast & Capacity Dashboard", layout="wide")
st.title("📊 Vendor Forecast vs Capacity Utilization")

# File uploads
booked_file = st.file_uploader("Upload Booked_Quantity.csv", type=["csv"])
forecast_file = st.file_uploader("Upload Forecast_JUNE_BW_converted.csv", type=["csv"])
capacity_file = st.file_uploader("Upload Vendor_allotted_capacity.csv", type=["csv"])

if booked_file and forecast_file and capacity_file:
    booked_df = pd.read_csv(booked_file)
    forecast_df = pd.read_csv(forecast_file)
    capacity_df = pd.read_csv(capacity_file)

    booked_df["PO exfac date"] = pd.to_datetime(booked_df["PO exfac date"], errors="coerce")
    booked_df["Month"] = booked_df["PO exfac date"].dt.to_period("M")
    booked_cleaned = booked_df.groupby(["VENDOR", "Month"], as_index=False).agg({"Qty": "sum"})

    # Auto-detect ex-factory column
    possible_date_cols = [col for col in forecast_df.columns if "ex-factory" in col.lower()]
    if not possible_date_cols:
        st.error("❌ Could not find a column like 'Vendor ex-factory' in your forecast file.")
        st.stop()
    forecast_df["Vendor ex-factory"] = pd.to_datetime(forecast_df[possible_date_cols[0]], errors="coerce")

    forecast_df["Month"] = forecast_df["Vendor ex-factory"].dt.to_period("M")
    forecast_cleaned = forecast_df.groupby(["Vendor Name", "Month"], as_index=False).agg(
        {"Confirmed New Planned Units": "sum"}
    ).rename(columns={"Vendor Name": "VENDOR"})

    combined_df = pd.merge(booked_cleaned, forecast_cleaned, on=["VENDOR", "Month"], how="outer")
    combined_df["Qty"] = combined_df["Qty"].fillna(0)
    combined_df["Confirmed New Planned Units"] = combined_df["Confirmed New Planned Units"].fillna(0)
    combined_df.rename(columns={"Qty": "Booked Qty", "Confirmed New Planned Units": "Forecast Qty"}, inplace=True)

    capacity_df["Vendor"] = capacity_df["Vendor"].str.strip().str.upper()
    capacity_months = [col for col in capacity_df.columns if "FM" in col and col.split()[0][:3] in calendar.month_abbr]
    capacity_melted = capacity_df.melt(id_vars=["Vendor"], value_vars=capacity_months, var_name="Month Name", value_name="Capacity")
    month_mapping = {month: f"{i:02}" for i, month in enumerate(calendar.month_abbr) if month}
    capacity_melted["Month"] = capacity_melted["Month Name"].str.extract(r'(\w{3})').iloc[:,0]
    capacity_melted["Month_Num"] = capacity_melted["Month"].map(month_mapping)
    capacity_melted["Year"] = 2025
    capacity_melted["Month"] = pd.to_datetime(capacity_melted["Year"].astype(str) + "-" + capacity_melted["Month_Num"]).dt.to_period("M")
    capacity_melted.drop(columns=["Month Name", "Month_Num", "Year"], inplace=True)
    capacity_cleaned = capacity_melted.groupby(["Vendor", "Month"], as_index=False).agg({"Capacity": "sum"})

    combined_df["VENDOR"] = combined_df["VENDOR"].str.strip().str.upper()
    final_df = pd.merge(combined_df, capacity_cleaned, left_on=["VENDOR", "Month"], right_on=["Vendor", "Month"], how="left")
    final_df.drop(columns=["Vendor"], inplace=True)
    final_df["Capacity"] = final_df["Capacity"].fillna(0)
    final_df["Utilization %"] = ((final_df["Booked Qty"] + final_df["Forecast Qty"]) / final_df["Capacity"]) * 100
    final_df["Utilization %"] = final_df["Utilization %"].replace([np.inf, -np.inf], np.nan).fillna(0).round(2)

    def get_flag(util):
        if util > 110:
            return "Overbooked"
        elif util < 70:
            return "Underutilized"
        else:
            return "Optimal"

    final_df["Flag"] = final_df["Utilization %"].apply(get_flag)

    vendors = final_df["VENDOR"].unique()
    months = sorted(final_df["Month"].unique())
    rows = []
    for vendor in vendors:
        for metric in ["CAPACITY", "BOOKEDQTY", "FORECAST QTY", "BALANCE CAPACITY", "Utilization % WITH COLOUR FLAG"]:
            row = {"Vendor / Month": vendor if metric == "CAPACITY" else "", "ITEM": metric}
            for month in months:
                record = final_df[(final_df["VENDOR"] == vendor) & (final_df["Month"] == month)]
                if not record.empty:
                    if metric == "CAPACITY":
                        value = record["Capacity"].values[0]
                    elif metric == "BOOKEDQTY":
                        value = record["Booked Qty"].values[0]
                    elif metric == "FORECAST QTY":
                        value = record["Forecast Qty"].values[0]
                    elif metric == "BALANCE CAPACITY":
                        value = record["Capacity"].values[0] - (record["Booked Qty"].values[0] + record["Forecast Qty"].values[0])
                    elif metric == "Utilization % WITH COLOUR FLAG":
                        value = f"{record['Utilization %'].values[0]}% - {record['Flag'].values[0]}"
                    else:
                        value = ""
                    row[month.strftime("'%b'%y")] = value
                else:
                    row[month.strftime("'%b'%y")] = ""
            rows.append(row)

    final_structured_df = pd.DataFrame(rows)
    st.success("✅ Report generated!")
    st.dataframe(final_structured_df, use_container_width=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_structured_df.to_excel(writer, index=False, sheet_name='Utilization_Report')
        workbook = writer.book
        worksheet = writer.sheets['Utilization_Report']
        red = workbook.add_format({'bg_color': '#FF0000'})
        yellow = workbook.add_format({'bg_color': '#FFFF00'})
        green = workbook.add_format({'bg_color': '#00FF00'})

        for row_idx, row in enumerate(final_structured_df.itertuples(index=False), start=1):
            if row.ITEM == "Utilization % WITH COLOUR FLAG":
                for col_idx in range(2, len(row)):
                    cell_val = row[col_idx]
                    if isinstance(cell_val, str):
                        if "Overbooked" in cell_val:
                            worksheet.write(row_idx, col_idx, cell_val, red)
                        elif "Underutilized" in cell_val:
                            worksheet.write(row_idx, col_idx, cell_val, yellow)
                        elif "Optimal" in cell_val:
                            worksheet.write(row_idx, col_idx, cell_val, green)

    st.download_button("📥 Download Excel Report", data=output.getvalue(), file_name="Vendor_Utilization_Report.xlsx")
