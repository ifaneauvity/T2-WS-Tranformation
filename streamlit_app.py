import streamlit as st
import pandas as pd
import re

# Streamlit app title
st.title("📊 WS Transformation")
st.write("Upload an Excel file and choose the transformation format.")

# Select transformation format
transformation_choice = st.radio("Select Transformation Format:", ["宏酒樽", "向日葵"])

if transformation_choice == "宏酒樽":
    raw_data_file = st.file_uploader("Upload Raw Sales Data", type=["xlsx"], key="new_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="new_mapping")
    
    if raw_data_file is not None and mapping_file is not None:
        # Find the sheet that contains "銷售(夜)" in the name
        xls = pd.ExcelFile(raw_data_file)
        sheet_name = next((sheet for sheet in xls.sheet_names if "銷售(夜)" in sheet), None)

        if sheet_name:
            df_raw = xls.parse(sheet_name)
            
            sheets_mapping = pd.ExcelFile(mapping_file).sheet_names  
            dfs_mapping = {sheet: pd.read_excel(mapping_file, sheet_name=sheet) for sheet in sheets_mapping}
            
            df_transformed = df_raw.iloc[:, [1, 2, 3, 4, 5, 6]].copy()
            df_transformed.columns = ["Date", "Outlet Code", "Outlet Name", "Product Code", "Product Name", "Number of Bottles"]
            
            # Add fixed columns
            df_transformed.insert(0, "Column1", "INV")
            df_transformed.insert(1, "Column2", "U")
            df_transformed.insert(2, "Column3", "30010085")
            df_transformed.insert(3, "Column4", "宏酒樽 ON")
            
            df_transformed["Date"] = pd.to_datetime(df_transformed["Date"]).dt.strftime('%Y%m%d')
            
            # Map product codes
            df_sku_mapping = dfs_mapping["SKU Mapping"]
            df_sku_mapping = df_sku_mapping[["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"]].drop_duplicates(subset="ASI_CRM_Offtake_Product__c")
            
            df_transformed = df_transformed.merge(
                df_sku_mapping,
                left_on="Product Code",
                right_on="ASI_CRM_Offtake_Product__c",
                how="left"
            )
            
            df_transformed.rename(columns={"ASI_CRM_SKU_Code__c": "SKU Code"}, inplace=True)
            df_transformed.drop(columns=["ASI_CRM_Offtake_Product__c"], inplace=True)
            
            # ✅ Fix Outlet Code Mapping Issue ✅
            df_transformed["Outlet Code"] = df_transformed["Outlet Code"].astype(str)

            # Optional replacement only if values are dates (skip if not needed)
            df_transformed["Outlet Code"] = df_transformed["Outlet Code"].replace({
                "2024-05-01 00:00:00": "5月1日",
                "2024-07-01 00:00:00": "7月1日",
                "2024-07-02 00:00:00": "07-02"
            })
            
            # Map customer codes
            df_customer_mapping = dfs_mapping["Customer Mapping"]
            df_customer_mapping = df_customer_mapping[["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"]].drop_duplicates(subset="ASI_CRM_Offtake_Customer_No__c")
            
            df_transformed = df_transformed.merge(
                df_customer_mapping,
                left_on="Outlet Code",
                right_on="ASI_CRM_Offtake_Customer_No__c",
                how="left"
            )
            
            df_transformed.rename(columns={"ASI_CRM_JDE_Cust_No_Formula__c": "PRT Customer Code"}, inplace=True)
            df_transformed.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "Outlet Code"], inplace=True)
            
            # Reorder the columns
            column_order = ["Column1", "Column2", "Column3", "Column4", "PRT Customer Code", "Outlet Name", "Date", "SKU Code", "Product Code", "Product Name", "Number of Bottles"]
            df_transformed = df_transformed[column_order]
            
            # ✅ Remove Exact Duplicates ✅
            duplicates = df_transformed[df_transformed.duplicated(keep=False)]

            if not duplicates.empty:
                st.warning("⚠️ Possible Duplicates Found:")
                st.dataframe(duplicates)
                
                # Drop duplicates
                df_transformed = df_transformed.drop_duplicates(keep='first')

            # Preview data in Streamlit
            st.write("✅ Processed Data Preview:")
            st.dataframe(df_transformed)
            
            # Export without headers
            output_filename = "processed_macro.xlsx"
            df_transformed.to_excel(output_filename, index=False, header=False)
            
            with open(output_filename, "rb") as f:
                st.download_button(label="📥 Download Processed File", data=f, file_name=output_filename)

elif transformation_choice == "向日葵":
    uploaded_file = st.file_uploader("Upload Raw Sales Data", type=["xlsx"], key="sunflower_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="sunflower_mapping")

    if uploaded_file is not None and mapping_file is not None:
        df = pd.read_excel(uploaded_file, header=None)

        # Create an empty list to store the extracted data
        data = []

        # Initialize variables to hold the current customer name, code, and date
        current_customer = None
        current_customer_code = None
        current_date = None

        # Start processing from row 8 (index 7)
        for i in range(7, len(df)):
            row = df.iloc[i]
            
            # Check if the row contains a customer name (by looking for "客戶名稱")
            if isinstance(row[0], str) and '客戶名稱' in row[0]:
                cleaned_text = re.sub(r'[\u200b\ufeff]', '', row[0]).strip()
                
                match = re.search(r'客戶編號[:：]\s*([\d\-]+).*客戶名稱[:：]\s*(.*)', cleaned_text)
                if match:
                    current_customer_code = match.group(1).strip()
                    current_customer = match.group(2).strip()
            
            # Check if the row contains a date
            if isinstance(row[0], str) and re.match(r'\d{3}/\d{2}/\d{2}', row[0]):
                year, month, day = map(int, row[0].split('/'))
                current_date = f'{year + 1911}{month:02}{day:02}'
            
            if pd.notna(row[1]):
                product_code = row[1]
                product_name = row[2]
                quantity = row[3]
                
                data.append([current_customer_code, current_customer, current_date, product_code, product_name, quantity])
        
        result_df = pd.DataFrame(data, columns=['Customer Code', 'Customer Name', 'Date', 'Product Code', 'Product Name', 'Quantity'])

        # ✅ Remove Exact Duplicates ✅
        duplicates = result_df[result_df.duplicated(keep=False)]

        if not duplicates.empty:
            st.warning("⚠️ Possible Duplicates Found:")
            st.dataframe(duplicates)
            result_df = result_df.drop_duplicates(keep='first')

        # Preview data in Streamlit
        st.write("✅ Processed Data Preview:")
        st.dataframe(result_df)

        output_filename = "processed_sunflower.xlsx"
        result_df.to_excel(output_filename, index=False, header=False)

        with open(output_filename, "rb") as f:
            st.download_button(label="📥 Download Processed File", data=f, file_name=output_filename)
