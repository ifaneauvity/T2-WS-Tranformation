import streamlit as st
import pandas as pd
import re

# Streamlit app title
st.title("📊 Excel Sales Data Processor")
st.write("Upload an Excel file and choose the transformation format.")

# Select transformation format
transformation_choice = st.radio("Select Transformation Format:", ["宏酒樽 Old Format", "宏酒樽 New Format"])

if transformation_choice == "宏酒樽 Old Format":
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"], key="old_format")
    
    if uploaded_file is not None:
        xls = pd.ExcelFile(uploaded_file)
        final_df = pd.DataFrame(columns=["Source Sheet", "Outlet", "Product", "Code", "Quantity", "Sales Date"])
        data = []
        
        for sheet in xls.sheet_names:
            df = xls.parse(sheet, header=None)
            start_row = 6
            outlets = df.iloc[start_row:, 1]
            products = df.iloc[start_row:, 2]
            quantities = df.iloc[start_row:, 4]
            sales_dates = df.iloc[start_row:, 3]
            valid_rows = outlets.notna()
            product_split = products[valid_rows].astype(str).str.extract(r'(\[.*?\]|【.*?】)?(.*)')
            
            sheet_data = pd.DataFrame({
                "Source Sheet": sheet,
                "Outlet": outlets[valid_rows].reset_index(drop=True),
                "Code": product_split[0].fillna("").reset_index(drop=True),
                "Product": product_split[1].str.strip().reset_index(drop=True),
                "Quantity": quantities[valid_rows].reset_index(drop=True),
                "Sales Date": sales_dates[valid_rows].reset_index(drop=True),
            })
            data.append(sheet_data)
        
        final_df = pd.concat(data, ignore_index=True)
        
        st.write("✅ Processed Data Preview:")
        st.dataframe(final_df)
        
        output_filename = "processed_old_format.xlsx"
        final_df.to_excel(output_filename, index=False, header=False)
        
        with open(output_filename, "rb") as f:
            st.download_button(label="📥 Download Processed File", data=f, file_name=output_filename)

elif transformation_choice == "宏酒樽 New Format":
    raw_data_file = st.file_uploader("Upload Raw Sales Data", type=["xlsx"], key="new_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="new_mapping")
    
    if raw_data_file is not None and mapping_file is not None:
        df_raw = pd.read_excel(raw_data_file, sheet_name=0)
        dfs_mapping = pd.read_excel(mapping_file, sheet_name=None)
        
        # Adjusted column selection (shifted right by 1 to correct misalignment)
        df_transformed = df_raw.iloc[:, [2, 3, 4, 5, 6, 7]].copy()
        df_transformed.columns = ["Date", "Outlet Code", "Outlet Name", "Product Code", "Product Name", "Number of Bottles"]
        
        df_transformed.insert(0, "Column1", "INV")
        df_transformed.insert(1, "Column2", "U")
        df_transformed.insert(2, "Column3", "30010085")
        df_transformed.insert(3, "Column4", "宏酒樽 ON")
        
        # Fix date formatting
        df_transformed["Date"] = pd.to_datetime(df_transformed["Date"], errors='coerce', infer_datetime_format=True)
        df_transformed["Date"] = df_transformed["Date"].dt.strftime('%Y%m%d')
        
        df_transformed["Outlet Code"] = df_transformed["Outlet Code"].astype(str).replace({
            "2024-05-01 00:00:00": "5月1日",
            "2024-07-01 00:00:00": "7月1日"
        })
        
        df_sku_mapping = dfs_mapping.get("SKU Mapping", pd.DataFrame())
        if not df_sku_mapping.empty:
            df_sku_mapping.columns = df_sku_mapping.columns.str.strip()
            st.write("Columns in SKU Mapping Sheet:", df_sku_mapping.columns.tolist())
            df_sku_mapping = df_sku_mapping.drop_duplicates(subset="ASI_CRM_Offtake_Product__c")
            if "ASI_CRM_Offtake_Product__c" in df_sku_mapping.columns and "ASI_CRM_SKU_Code__c" in df_sku_mapping.columns:
                df_transformed = df_transformed.merge(
                    df_sku_mapping[["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"]],
                    left_on="Product Code", right_on="ASI_CRM_Offtake_Product__c", how="left"
                )
                df_transformed.rename(columns={"ASI_CRM_SKU_Code__c": "SKU Code"}, inplace=True)
                df_transformed.drop(columns=["ASI_CRM_Offtake_Product__c"], inplace=True, errors='ignore')
            else:
                st.error("❌ SKU Mapping sheet is missing required columns.")
        
        df_customer_mapping = dfs_mapping.get("Customer Mapping", pd.DataFrame())
        if not df_customer_mapping.empty:
            df_customer_mapping.columns = df_customer_mapping.columns.str.strip()
            st.write("Columns in Customer Mapping Sheet:", df_customer_mapping.columns.tolist())
            df_customer_mapping = df_customer_mapping.drop_duplicates(subset="ASI_CRM_Offtake_Customer_No__c")
            if "ASI_CRM_Offtake_Customer_No__c" in df_customer_mapping.columns and "ASI_CRM_JDE_Cust_No_Formula__c" in df_customer_mapping.columns:
                df_transformed = df_transformed.merge(
                    df_customer_mapping[["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"]],
                    left_on="Outlet Code", right_on="ASI_CRM_Offtake_Customer_No__c", how="left"
                )
                df_transformed.rename(columns={"ASI_CRM_JDE_Cust_No_Formula__c": "PRT Customer Code"}, inplace=True)
                df_transformed.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "Outlet Code"], inplace=True, errors='ignore')
            else:
                st.error("❌ Customer Mapping sheet is missing required columns.")
        
        column_order = ["Column1", "Column2", "Column3", "Column4", "PRT Customer Code", "Outlet Name", "Date", "SKU Code", "Product Code", "Product Name", "Number of Bottles"]
        df_transformed = df_transformed[column_order]
        
        output_filename = "processed_new_format.xlsx"
        df_transformed.to_excel(output_filename, index=False, header=False)
        
        st.write("✅ Processed Data Preview:")
        st.dataframe(df_transformed)
        
        with open(output_filename, "rb") as f:
            st.download_button(label="📥 Download Processed File", data=f, file_name=output_filename)
