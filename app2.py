import pandas as pd
import glob
import os
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import numpy as np
import streamlit as st

# Streamlit app title
st.title("IRS Report Processor + Tolerance Break Trend Analytics")

# File uploader for multiple Excel files
uploaded_files = st.file_uploader(
    "Drag and drop or select NAV reports (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.warning("Please upload at least one NAV report (.xlsx) to continue.")
    st.stop()

st.write(f"Uploaded {len(uploaded_files)} files for processing.")

# Initialize an empty list to store DataFrames
all_data = []

# Process each uploaded file, focusing only on the "IRS" sheet
for uploaded_file in uploaded_files:
    try:
        # Read only the "IRS" sheet
        raw_data = pd.read_excel(uploaded_file, sheet_name="IRS", header=[12, 13])

        # Generate column names by merging row 13 and 14
        new_columns = [
            f"{str(upper).strip()}_{str(lower).strip()}" if str(upper) != 'nan' else str(lower).strip()
            for upper, lower in zip(raw_data.columns.get_level_values(0), raw_data.columns.get_level_values(1))
        ]
        raw_data.columns = new_columns

        # Drop potential blank rows
        raw_data = raw_data.iloc[1:].reset_index(drop=True)

        # Add a Report Date column
        report_date = uploaded_file.name.split('-')[-1].split('.')[0] if '-' in uploaded_file.name else "Unknown Date"
        raw_data["Report Date"] = report_date

        # Append processed data
        all_data.append(raw_data)

    except Exception as e:
        st.error(f"Error processing {uploaded_file.name}: {e}")
        continue

# Combine all DataFrames
if all_data:
    data = pd.concat(all_data, ignore_index=True)
    st.success("All files processed successfully!")
else:
    st.error("No valid data found after processing.")
    st.stop()

# Filter the DataFrame to only keep relevant columns
filtered_columns = {
    "GTID_Unnamed: 0_level_1": "Trade ID 1",
    "Original GTID_Unnamed: 1_level_1": "Original GTID",
    "Counterparty / Clearing Member_MV Base": "Counterparty MV Base",
    "SS&C GlobeOp Source_MV Base": "SS&C MV Base",
    "Instrument Sub Type_Unnamed: 6_level_1": "Product Sub Type",
    "SS&C GlobeOp Source_IR DV01": "SS&C IR DV01",
    "SS&C GlobeOp Trade Attributes_Trade Date": "Trade Date",
    "SS&C GlobeOp Trade Attributes_Effective Date": "Effective Date",
    "SS&C GlobeOp Trade Attributes_Maturity Date": "Maturity Date",
    "SS&C GlobeOp Trade Attributes_Ccy": "Currency",
    "SS&C GlobeOp Trade Attributes_Notional": "Notional",
    "SS&C GlobeOp Trade Attributes_Rec Rate": "Rec Rate",
    "SS&C GlobeOp Trade Attributes_Pay Rate": "Pay Rate",
    "Counterparty/ Clearing Member_Final Source Load Time": "Final Source Load Time",
    "Final Source vs Counterparty / Clearing Member_Difference in MV": "Difference in MV",
    "Final Source vs Counterparty / Clearing Member_NAV Tolerance Analysis": "NAV Tolerance Analysis",
    "Final Source vs Counterparty / Clearing Member_Diff. in MV/IR DV01 or Diff. in MV/IDV01": "Diff. in MV/IR DV01"
}
filtered_data = data[list(filtered_columns.keys())].rename(columns=filtered_columns)

filtered_data = filtered_data.dropna(subset=["Trade ID 1"])

# Convert Final Source Load Time to DDMMYYYY format
filtered_data['Final Source Load Time'] = pd.to_datetime(filtered_data['Final Source Load Time']).dt.strftime('%d%m%Y')

# Set up new columns
filtered_data['Sensitivity Breach'] = None
filtered_data['Tolerance Breach'] = None
filtered_data['Urgent Escalation Required'] = None

# Ask the user for the client name
client = st.text_input("Please enter the client you are analyzing (e.g., ASGARD): ").strip()

# Only proceed if the client is provided
if client:
    st.write(f"Processing data for client: **{client}**")

    if client.upper() == "ASGARD":
        # Apply conditions for Tolerance Breach
        filtered_data['Tolerance Breach'] = filtered_data['NAV Tolerance Analysis'].abs() > 1

    # Apply Sensitivity Breach logic to ASGARD
    ccy_list = ['USD', 'CAD', 'JPY', 'AUD', 'NZD', 'GBP', 'EUR', 'CHF', 'SEK', 'NOK']
    filtered_data['Sensitivity Breach'] = filtered_data.apply(
        lambda row: (
            "TRUE" if (
                row['Currency'] in ccy_list and
                (
                    (row['Product Sub Type'] == "Plain Vanilla" and abs(row['Diff. in MV/IR DV01']) > 2) or
                    (row['Product Sub Type'] == "OIS" and abs(row['Diff. in MV/IR DV01']) > 1) or
                    (row['Product Sub Type'] == "MTM Cross Currency Swap" and abs(row['Diff. in MV/IR DV01']) > 4)
                )
            ) else (
                "TRUE" if (
                    row['Currency'] not in ccy_list and
                    (
                        (row['Product Sub Type'] == "Plain Vanilla" and abs(row['Diff. in MV/IR DV01']) > 5) or
                        (row['Product Sub Type'] == "OIS" and abs(row['Diff. in MV/IR DV01']) > 8) or
                        (row['Product Sub Type'] == "MTM Cross Currency Swap" and abs(row['Diff. in MV/IR DV01']) > 9)
                    )
                ) else "FALSE"
            )
        ), axis=1
    )

    # Create index column
    filtered_data["Index"] = None
    filtered_data["Index_Maturity"] = None

    def get_index(row):
        def is_numeric_rate(value):
            if not isinstance(value, str):
                return True
            cleaned = value.replace('%', '').replace(' ', '')
            return cleaned.replace('.', '', 1).replace('-', '', 1).isdigit()

        rec_rate = row['Rec Rate']
        pay_rate = row['Pay Rate']

        if not is_numeric_rate(rec_rate):
            return rec_rate
        elif not is_numeric_rate(pay_rate):
            return pay_rate
        return None

    filtered_data['Index'] = filtered_data.apply(get_index, axis=1)

    # Extract the year from Maturity Date and add a "Maturity Year" column
    filtered_data['Maturity Year'] = pd.to_datetime(filtered_data['Maturity Date'], errors='coerce').dt.year.astype('Int64')

    # Create a new column for the Index and Maturity Year concatenation
    filtered_data['Index_Maturity'] = filtered_data['Index'] + "_" + filtered_data['Maturity Year'].astype(str)

    st.write(filtered_data)

    # Save the updated DataFrame to an Excel file
    output_file = "Processed_ASGARD_Report_with_Breaches.xlsx"
    filtered_data.to_excel(output_file, index=False, engine='openpyxl', sheet_name="Processed Report")

    # Apply conditional formatting for TRUE/FALSE
    wb = load_workbook(output_file)
    ws = wb["Processed Report"]
    filtered_data['Sensitivity Breach'] = filtered_data['Sensitivity Breach'].astype(str)
    filtered_data['Tolerance Breach'] = filtered_data['Tolerance Breach'].astype(str)

    columns_to_format = ["Sensitivity Breach", "Tolerance Breach"]
    true_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    false_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    for col in columns_to_format:
        if col in filtered_data.columns:
            col_idx = filtered_data.columns.get_loc(col) + 1
            for row in range(2, len(filtered_data) + 2):
                cell = ws.cell(row=row, column=col_idx)
                if str(cell.value).upper() == "TRUE":
                    cell.fill = true_fill
                elif str(cell.value).upper() == "FALSE":
                    cell.fill = false_fill
        else:
            st.warning(f"Warning: Column '{col}' not found in DataFrame. Skipping conditional formatting.")

    wb.save(output_file)
    st.write(f"Conditional formatting applied and saved to: {output_file}")

    # Ensure Tolerance Breach is a boolean before plotting
    filtered_data['Tolerance Breach'] = filtered_data['Tolerance Breach'].str.upper() == "TRUE"

    # Visualization - Breaches grouped by Product Type and Ccy
    filtered_data['Product_Ccy'] = filtered_data['Product Sub Type'] + "_" + filtered_data['Currency']

    sensitivity_breach_counts = filtered_data[filtered_data['Sensitivity Breach'] == "TRUE"]['Product_Ccy'].value_counts()
    tolerance_breach_counts = filtered_data[filtered_data['Tolerance Breach'] == True]['Product_Ccy'].value_counts()

    all_product_ccy = sensitivity_breach_counts.index.union(tolerance_breach_counts.index)
    sensitivity_breach_counts = sensitivity_breach_counts.reindex(all_product_ccy, fill_value=0)
    tolerance_breach_counts = tolerance_breach_counts.reindex(all_product_ccy, fill_value=0)

    # Plotting both charts stacked
    fig, axes = plt.subplots(3, 1, figsize=(14, 16), sharex=True)

    def add_bar_labels(ax):
        for bar in ax.patches:
            height = bar.get_height()
            if height > 0:
                ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    height * 0.5,
                    str(int(height)),
                    ha='center',
                    va='center',
                    fontsize=12,
                    fontweight='bold',
                    color='white'
                )

    # Sensitivity Breach Chart
    bars1 = axes[0].bar(sensitivity_breach_counts.index, sensitivity_breach_counts, color='skyblue', edgecolor='black')
    axes[0].set_title("Count of Sensitivity Breaches by Product Type and Ccy", fontsize=14)
    axes[0].set_ylabel("Count of Sensitivity Breaches", fontsize=12)
    axes[0].tick_params(axis='x', rotation=45, labelsize=10)
    axes[0].grid(True, linestyle='--', alpha=0.6)
    add_bar_labels(axes[0])

    # Tolerance Breach Chart
    bars2 = axes[1].bar(tolerance_breach_counts.index, tolerance_breach_counts, color='lightcoral', edgecolor='black')
    axes[1].set_title("Count of Tolerance Breaches by Product Type and Ccy", fontsize=14)
    axes[1].set_ylabel("Count of Tolerance Breaches", fontsize=12)
    axes[1].tick_params(axis='x', rotation=45, labelsize=10)
    axes[1].grid(True, linestyle='--', alpha=0.6)
    add_bar_labels(axes[1])

    # Immediate Attention Required Chart
    immediate_attention_df = filtered_data[
        (filtered_data['Sensitivity Breach'] == "TRUE") &
        (filtered_data['Tolerance Breach'] == True)
    ]
    immediate_attention_counts = immediate_attention_df['Product_Ccy'].value_counts()
    bars3 = axes[2].bar(immediate_attention_counts.index, immediate_attention_counts, color='darkred', edgecolor='black')
    axes[2].set_title("Breaks Requiring Immediate Attention (Both Sensitivity & Tolerance Breaches)", fontsize=14)
    axes[2].set_ylabel("Count of Critical Breaches", fontsize=12)
    axes[2].tick_params(axis='x', rotation=45, labelsize=10)
    axes[2].grid(True, linestyle='--', alpha=0.6)
    add_bar_labels(axes[2])

    plt.tight_layout()
    st.pyplot(fig)

    # Trend Analysis
    filtered_df = filtered_data[
        (filtered_data['Sensitivity Breach'] == "TRUE") &
        (filtered_data['Product Sub Type'] != "MTM Cross Currency Swap")
    ]
    breach_counts = filtered_df.groupby(['Currency', 'Index_Maturity']).size().reset_index(name='Breach Count')
    top_5_per_currency = breach_counts.groupby('Currency').apply(lambda x: x.nlargest(5, 'Breach Count')).reset_index(drop=True)

    filtered_data['Plot Date'] = pd.to_datetime(filtered_data['Final Source Load Time'], format='%d%m%Y')

    unique_ccys = top_5_per_currency['Currency'].unique()
    fig, axes = plt.subplots(len(unique_ccys), 1, figsize=(15, 5 * len(unique_ccys)), sharex=True)

    if len(unique_ccys) == 1:
        axes = [axes]

    for ax, ccy in zip(axes, unique_ccys):
        top_5_indices = top_5_per_currency[top_5_per_currency['Currency'] == ccy]['Index_Maturity'].tolist()
        ccy_data = filtered_data[
            (filtered_data['Currency'] == ccy) &
            (filtered_data['Index_Maturity'].isin(top_5_indices)) &
            (filtered_data['Sensitivity Breach'] == "TRUE") &
            (filtered_data['Product Sub Type'] != "MTM Cross Currency Swap")
        ].groupby(['Plot Date', 'Index_Maturity']).size().unstack(fill_value=0)

        for column in ccy_data.columns:
            ax.plot(ccy_data.index, ccy_data[column], marker='o', label=column)

        ax.set_title(f"Daily Sensitivity Breaches for {ccy} (Top 5 Index Maturities)", fontsize=14)
        ax.set_xlabel("Date", fontsize=12)
        ax.set_ylabel("Number of Breaches", fontsize=12)
        ax.grid(True, linestyle='--', alpha=0.6)
        ax.legend(title="Index Maturity", fontsize=8, bbox_to_anchor=(1.05, 1), loc='upper left')

    plt.xticks(rotation=45)
    plt.tight_layout()
    st.pyplot(fig)

    # Provide download link for the processed Excel file
    st.subheader("Download Processed Data")
    st.write("Click the button below to download the processed Excel file.")
    with open(output_file, "rb") as file:
        btn = st.download_button(
            label="Download Excel",
            data=file,
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )