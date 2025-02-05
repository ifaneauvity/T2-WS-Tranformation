import streamlit as st
import pandas as pd
import re

# Streamlit app title
st.title("📊 Excel Sales Data Processor")
st.write("Upload an Excel file, and this app will clean and organize the data automatically.")

# File uploader
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Load the Excel file
    xls = pd.ExcelFile(uploaded_file)

    # Create an empty DataFrame for storing cleaned data
    final_df = pd.DataFrame(columns=["Source Sheet", "Outlet", "Product", "Code", "Quantity", "Sales Date"])

    # Process each sheet in the Excel file
    data = []
    for sheet in xls.sheet_names:
        df = xls.parse(sheet, header=None)  # Load without headers
        start_row = 6  # Data starts at row 7 in Excel (index 6 in pandas)
        
        # Extract relevant columns
        outlets = df.iloc[start_row:, 1]  # Column B - Outlet
        products = df.iloc[start_row:, 2]  # Column C - Product
        quantities = df.iloc[start_row:, 4]  # Column E - Quantity
        sales_dates = df.iloc[start_row:, 3]  # Column D - Sales Date

        # Remove empty rows based on Outlet (ensures all columns align)
        valid_rows = outlets.notna()

        # Extract product code inside [],【】, if available
        product_split = products[valid_rows].astype(str).str.extract(r'(\[.*?\]|【.*?】)?(.*)')

        # Create DataFrame for cleaned data
        sheet_data = pd.DataFrame({
            "Source Sheet": sheet,
            "Outlet": outlets[valid_rows].reset_index(drop=True),
            "Code": product_split[0].fillna("").reset_index(drop=True),  # Product Code
            "Product": product_split[1].str.strip().reset_index(drop=True),  # Product Name
            "Quantity": quantities[valid_rows].reset_index(drop=True),
            "Sales Date": sales_dates[valid_rows].reset_index(drop=True),
        })

        # Append to final dataset
        data.append(sheet_data)

    # Convert list to final DataFrame
    final_df = pd.concat(data, ignore_index=True)

    # Display processed data
    st.write("✅ Processed Data Preview:")
    st.dataframe(final_df)

    # Save separate files based on Sales Date
    unique_dates = final_df["Sales Date"].unique()
    output_files = {}

    for date in unique_dates:
        date_df = final_df[final_df["Sales Date"] == date]

        # Replace invalid filename characters
        safe_date = re.sub(r'[\/:*?"<>|]', '_', str(date))
        filename = f"{safe_date}.xlsx"

        date_df.to_excel(filename, index=False)
        output_files[date] = filename

    # Download processed files
    st.write("📥 Download Processed Files:")
    for date, file in output_files.items():
        with open(file, "rb") as f:
            st.download_button(label=f"Download {file}", data=f, file_name=file)
