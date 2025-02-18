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
    "Drag and drop or select ASGARD reports (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.warning("Please upload at least one ASGARD report (.xlsx) to continue.")
    st.stop()

st.write(f"Uploaded {len(uploaded_files)} files for processing.")


# Function to extract date from filename (if necessary)
def extract_date_from_filename(filename):
    return filename.split('-')[-1].split('.')[0] if '-' in filename else "Unknown Date"


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
        report_date = extract_date_from_filename(uploaded_file.name)
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
# Step 12: Ask the user which client they are analyzing
client = st.text_input("Please enter the client you are analyzing (e.g., ASGARD): ").strip()

filtered_columns = {
    "GTID_Unnamed: 0_level_1": "Trade ID 1",  # Renaming this column
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

# Filter the DataFrame to only keep these columns
filtered_data = data[list(filtered_columns.keys())].rename(columns=filtered_columns)


# Convert Final Source Load Time to DDMMYYYY format
filtered_data['Final Source Load Time'] = pd.to_datetime(filtered_data['Final Source Load Time']).dt.strftime('%d%m%Y')

#  Set up new columns
filtered_data['Sensitivity Breach'] = None
filtered_data['Tolerance Breach'] = None
filtered_data['Urgent Escalation Required'] = None



if client.upper() == "ASGARD":
    # Apply conditions for Tolerance Breach
    filtered_data['Tolerance Breach'] = filtered_data['NAV Tolerance Analysis'].abs() > 1

# Apply Sensitivity Breach logic to ASGARD

# Define the currency list
ccy_list = ['USD', 'CAD', 'JPY', 'AUD', 'NZD', 'GBP', 'EUR', 'CHF', 'SEK', 'NOK']

filtered_data['Sensitivity Breach'] = filtered_data.apply(
    lambda row: (
        "TRUE" if (
            # For currencies in the list
            row['Currency'] in ccy_list and
            (
                (row['Product Sub Type'] == "Plain Vanilla" and abs(row['Diff. in MV/IR DV01']) > 2) or
                (row['Product Sub Type'] == "OIS" and abs(row['Diff. in MV/IR DV01']) > 1) or
                (row['Product Sub Type'] == "MTM Cross Currency Swap" and abs(row['Diff. in MV/IR DV01']) > 4)
            )
        ) else (
            # For currencies outside the list
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


#Create index column

filtered_data["Index"] = None
filtered_data["Index_Maturity"] = None

def get_index(row):
    # Function to check if a value is a pure number (including negative numbers and percentages)
    def is_numeric_rate(value):
        if not isinstance(value, str):
            return True  # Already a number (float/int)
        # Remove '%', spaces, and check if it can be converted to a float
        cleaned = value.replace('%', '').replace(' ', '')
        return cleaned.replace('.', '', 1).replace('-', '', 1).isdigit()

    rec_rate = row['Rec Rate']
    pay_rate = row['Pay Rate']

    # Check if Rec Rate is NOT numeric, meaning it's likely an index
    if not is_numeric_rate(rec_rate):
        return rec_rate  # Rec Rate is a text (e.g., an index like "DKKCIBOR6M")
    # If Rec Rate is numeric, check Pay Rate
    elif not is_numeric_rate(pay_rate):
        return pay_rate  # Pay Rate is- a text (index)
    # If both are numeric, return None
    return None

# Apply the function to populate the Index column
filtered_data['Index'] = filtered_data.apply(get_index, axis=1)

# Optional: Verify the Index column doesn't contain any numeric values
numeric_indices = filtered_data['Index'].str.contains(r'^[\d\.]+%?$', na=False)
if numeric_indices.any():
    st.warning("Warning: Some numeric values found in Index column")
    st.write(filtered_data[numeric_indices][['Rec Rate', 'Pay Rate', 'Index']])

# Extract the year from Maturity Date and add a "Maturity Year" column
filtered_data['Maturity Year'] = pd.to_datetime(filtered_data['Maturity Date'], errors='coerce').dt.year.astype('Int64')

# Create a new column for the Index and Maturity Year concatenation
filtered_data['Index_Maturity'] = filtered_data['Index'] + "_" + filtered_data['Maturity Year'].astype(str)

st.write(filtered_data)

# Step 11: Save the updated DataFrame to an Excel file
output_file = "Processed_ASGARD_Report_with_Breaches.xlsx"
filtered_data.to_excel(output_file, index=False, engine='openpyxl', sheet_name="Processed Report")

# Step 14: Apply conditional formatting for TRUE/FALSE
wb = load_workbook(output_file)

ws = wb["Processed Report"]
# Ensure both columns are in string format for consistent conditional formatting
filtered_data['Sensitivity Breach'] = filtered_data['Sensitivity Breach'].astype(str)
filtered_data['Tolerance Breach'] = filtered_data['Tolerance Breach'].astype(str)

# Identify columns for conditional formatting
columns_to_format = ["Sensitivity Breach", "Tolerance Breach"]
true_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # Red for TRUE
false_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Green for FALSE

# Apply conditional formatting for both columns
for col in columns_to_format:
    if col in filtered_data.columns:  # Ensure the column exists
        col_idx = filtered_data.columns.get_loc(col) + 1  # Get column index (1-based for Excel)
        for row in range(2, len(filtered_data) + 2):  # Start from row 2 (skip header)
            cell = ws.cell(row=row, column=col_idx)
            if str(cell.value).upper() == "TRUE":
                cell.fill = true_fill
            elif str(cell.value).upper() == "FALSE":
                cell.fill = false_fill
    else:
        st.warning(f"Warning: Column '{col}' not found in DataFrame. Skipping conditional formatting.")

# Save the workbook with formatting
wb.save(output_file)
st.write(f"Conditional formatting applied and saved to: {output_file}")

# Ensure Tolerance Breach is a boolean before plotting
filtered_data['Tolerance Breach'] = filtered_data['Tolerance Breach'].str.upper() == "TRUE"

# Step 14: Visualization - Breaches grouped by Product Type and Ccy
filtered_data['Product_Ccy'] = filtered_data['Product Sub Type'] + "_" + filtered_data['Currency']  # Combine Product Type and Ccy

# Count Sensitivity Breach (TRUE) grouped by Product_Ccy
sensitivity_breach_counts = filtered_data[filtered_data['Sensitivity Breach'] == "TRUE"]['Product_Ccy'].value_counts()

# Count Tolerance Breach (TRUE) grouped by Product_Ccy
tolerance_breach_counts = filtered_data[filtered_data['Tolerance Breach'] == True]['Product_Ccy'].value_counts()

# Align indices of both counts (fill missing values with 0)
all_product_ccy = sensitivity_breach_counts.index.union(tolerance_breach_counts.index)
sensitivity_breach_counts = sensitivity_breach_counts.reindex(all_product_ccy, fill_value=0)
tolerance_breach_counts = tolerance_breach_counts.reindex(all_product_ccy, fill_value=0)

# Plotting both charts stacked
fig, axes = plt.subplots(3, 1, figsize=(14, 16), sharex=True)

# Function to add values inside bars
def add_bar_labels(ax):
    for bar in ax.patches:
        height = bar.get_height()
        if height > 0:
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                height * 0.5,  # Position inside the bar
                str(int(height)),
                ha='center',
                va='center',  # Centered inside the bar
                fontsize=12,
                fontweight='bold',
                color='white'  # White text for contrast
            )

# Sensitivity Breach Chart
bars1 = axes[0].bar(
    sensitivity_breach_counts.index,
    sensitivity_breach_counts,
    color='skyblue',
    edgecolor='black'  # Solid border
)
axes[0].set_title("Count of Sensitivity Breaches by Product Type and Ccy", fontsize=14)
axes[0].set_ylabel("Count of Sensitivity Breaches", fontsize=12)
axes[0].tick_params(axis='x', rotation=45, labelsize=10)
axes[0].grid(True, linestyle='--', alpha=0.6)  # Grid background
add_bar_labels(axes[0])  # Add labels inside bars

# Tolerance Breach Chart
bars2 = axes[1].bar(
    tolerance_breach_counts.index,
    tolerance_breach_counts,
    color='lightcoral',
    edgecolor='black'  # Solid border
)
axes[1].set_title("Count of Tolerance Breaches by Product Type and Ccy", fontsize=14)
axes[1].set_ylabel("Count of Tolerance Breaches", fontsize=12)
axes[1].tick_params(axis='x', rotation=45, labelsize=10)
axes[1].grid(True, linestyle='--', alpha=0.6)  # Grid background
add_bar_labels(axes[1])  # Add labels inside bars

# New third chart (Immediate Attention Required)
# Filter for trades with both Sensitivity and Tolerance breaches
immediate_attention_df = filtered_data[
    (filtered_data['Sensitivity Breach'] == "TRUE") &
    (filtered_data['Tolerance Breach'] == True)
]
immediate_attention_counts = immediate_attention_df['Product_Ccy'].value_counts()

bars3 = axes[2].bar(
    immediate_attention_counts.index,
    immediate_attention_counts,
    color='darkred',  # Darker red to indicate urgency
    edgecolor='black'
)
axes[2].set_title("Breaks Requiring Immediate Attention (Both Sensitivity & Tolerance Breaches)", fontsize=14)
axes[2].set_ylabel("Count of Critical Breaches", fontsize=12)
axes[2].tick_params(axis='x', rotation=45, labelsize=10)
axes[2].grid(True, linestyle='--', alpha=0.6)
add_bar_labels(axes[2])

# Adjust layout for all three charts
plt.tight_layout()

# Display the plots in Streamlit
st.pyplot(fig)

# Trend Analysis

filtered_df = filtered_data[
    (filtered_data['Sensitivity Breach'] == "TRUE") &
    (filtered_data['Product Sub Type'] != "MTM Cross Currency Swap")
]

# Group by Index_Maturity and count Sensitivity Breaches
breach_counts = filtered_df.groupby(['Currency', 'Index_Maturity']).size().reset_index(name='Breach Count')

# Get the top 5 Index Maturities **per currency**
top_5_per_currency = breach_counts.groupby('Currency').apply(lambda x: x.nlargest(5, 'Breach Count')).reset_index(drop=True)

# Convert Final Source Load Time to datetime for plotting
filtered_data['Plot Date'] = pd.to_datetime(filtered_data['Final Source Load Time'], format='%d%m%Y')

# Create a figure with subplots for each unique currency
unique_ccys = top_5_per_currency['Currency'].unique()
fig, axes = plt.subplots(len(unique_ccys), 1, figsize=(15, 5 * len(unique_ccys)), sharex=True)

# Ensure axes is iterable even if there's only one currency
if len(unique_ccys) == 1:
    axes = [axes]

# Plot data for each currency separately
for ax, ccy in zip(axes, unique_ccys):
    top_5_indices = top_5_per_currency[top_5_per_currency['Currency'] == ccy]['Index_Maturity'].tolist()

    # Filter dataset for the selected currency and top 5 Index_Maturities
    ccy_data = filtered_data[
        (filtered_data['Currency'] == ccy) &
        (filtered_data['Index_Maturity'].isin(top_5_indices)) &
        (filtered_data['Sensitivity Breach'] == "TRUE") &
        (filtered_data['Product Sub Type'] != "MTM Cross Currency Swap")
    ].groupby(['Plot Date', 'Index_Maturity']).size().unstack(fill_value=0)

    # Plot each Index Maturity in this currency's chart
    for column in ccy_data.columns:
        ax.plot(ccy_data.index, ccy_data[column], marker='o', label=column)  # Dots added back

        # Customize each subplot
    ax.set_title(f"Daily Sensitivity Breaches for {ccy} (Top 5 Index Maturities)", fontsize=14)
    ax.set_xlabel("Date", fontsize=12)
    ax.set_ylabel("Number of Breaches", fontsize=12)
    ax.grid(True, linestyle='--', alpha=0.6)
    ax.legend(title="Index Maturity", fontsize=8, bbox_to_anchor=(1.05, 1), loc='upper left')

# Rotate x-axis labels for readability
plt.xticks(rotation=45)

# Adjust layout to prevent overlap
plt.tight_layout()

# Show all plots
st.pyplot(fig)


# Step 14: Provide download link for the processed Excel file
st.subheader("Download Processed Data")
st.write("Click the button below to download the processed Excel file.")
with open(output_file, "rb") as file:
    btn = st.download_button(
        label="Download Excel",
        data=file,
        file_name=output_file,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

