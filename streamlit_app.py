import streamlit as st
import pandas as pd
import re

# Streamlit app title
st.title("📊 WS Transformation")
st.write("Upload an Excel file and choose the transformation format.")

# Select transformation format
transformation_choice = st.radio("Select Transformation Format:", ["30010085 宏酒樽 (夜)", "30010203 宏酒樽 (日)", "30010061 向日葵", "30010010 酒倉盛豐行", "30010013 酒田", "30010059 誠邦有限公司"])

if transformation_choice == "30010085 宏酒樽 (夜)":
    raw_data_file = st.file_uploader("Upload Raw Sales Data", type=["xlsx"], key="new_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="new_mapping")
    
    if raw_data_file is not None and mapping_file is not None:
        # Find the sheet that contains "夜" in the name
        xls = pd.ExcelFile(raw_data_file)
        sheet_name = next((sheet for sheet in xls.sheet_names if "夜" in sheet), None)

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
            
            # ✅🔄 Updated Customer Mapping with 30010085 Filter
            df_customer_mapping = dfs_mapping["Customer Mapping"]
            df_customer_mapping = df_customer_mapping[
                df_customer_mapping["ASI_CRM_Mapping_Cust_No__c"] == 30010085
            ][["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"]].drop_duplicates(
                subset="ASI_CRM_Offtake_Customer_No__c"
            )
            
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

            # Preview data in Streamlit
            st.write("✅ Processed Data Preview:")
            st.dataframe(df_transformed)
            
            # Export without headers
            output_filename = "processed_macro.xlsx"
            df_transformed.to_excel(output_filename, index=False, header=False)
            
            with open(output_filename, "rb") as f:
                st.download_button(label="📥 Download Processed File", data=f, file_name=output_filename)

elif transformation_choice == "30010203 宏酒樽 (日)":
    raw_data_file = st.file_uploader("Upload Raw Sales Data", type=["xlsx"], key="new_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="new_mapping")
    
    if raw_data_file is not None and mapping_file is not None:
        # Find the sheet that contains "日" in the name
        xls = pd.ExcelFile(raw_data_file)
        sheet_name = next((sheet for sheet in xls.sheet_names if "日" in sheet), None)

        if sheet_name:
            df_raw = xls.parse(sheet_name)
            
            sheets_mapping = pd.ExcelFile(mapping_file).sheet_names  
            dfs_mapping = {sheet: pd.read_excel(mapping_file, sheet_name=sheet) for sheet in sheets_mapping}
            
            df_transformed = df_raw.iloc[:, [1, 2, 3, 4, 5, 6]].copy()
            df_transformed.columns = ["Date", "Outlet Code", "Outlet Name", "Product Code", "Product Name", "Number of Bottles"]
            
            # Add fixed columns
            df_transformed.insert(0, "Column1", "INV")
            df_transformed.insert(1, "Column2", "U")
            df_transformed.insert(2, "Column3", "30010203")
            df_transformed.insert(3, "Column4", "宏酒樽 OFF")
            
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
            
            # ✅🔄 Updated Customer Mapping with 30010085 Filter
            df_customer_mapping = dfs_mapping["Customer Mapping"]
            df_customer_mapping = df_customer_mapping[
                df_customer_mapping["ASI_CRM_Mapping_Cust_No__c"] == 30010203
            ][["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"]].drop_duplicates(
                subset="ASI_CRM_Offtake_Customer_No__c"
            )
            
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

            # Preview data in Streamlit
            st.write("✅ Processed Data Preview:")
            st.dataframe(df_transformed)
            
            # Export without headers
            output_filename = "processed_macro.xlsx"
            df_transformed.to_excel(output_filename, index=False, header=False)
            
            with open(output_filename, "rb") as f:
                st.download_button(label="📥 Download Processed File", data=f, file_name=output_filename)

elif transformation_choice == "30010061 向日葵":
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

        # Add fixed columns
        result_df.insert(0, 'Column1', 'INV')
        result_df.insert(1, 'Column2', 'U')
        result_df.insert(2, 'Column3', '30010061')
        result_df.insert(3, 'Column4', '向日葵')

        # ✅ Remove exact duplicates
        result_df.drop_duplicates(keep='first', inplace=True)

        # Preview data in Streamlit
        st.write("✅ Processed Data Preview:")
        st.dataframe(result_df)

        output_filename = "processed_sunflower.xlsx"
        result_df.to_excel(output_filename, index=False, header=False)

        with open(output_filename, "rb") as f:
            st.download_button(label="📥 Download Processed File", data=f, file_name=output_filename)

elif transformation_choice == "30010010 酒倉盛豐行":
    raw_data_file = st.file_uploader("Upload Raw Sales Data", type=["xlsx"], key="sakakura_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="sakakura_mapping")

    if raw_data_file and mapping_file:
        raw_df = pd.read_excel(raw_data_file, sheet_name="給廠商", header=None)

        # Extract date from cell A5
        date_string = str(raw_df.iloc[4, 0])
        match = re.search(r'至\s*(\d{3}/\d{2}/\d{2})', date_string)
        if match:
            roc_date = match.group(1)
            year, month, day = map(int, roc_date.split('/'))
            final_date = f"{year + 1911}{month:02d}{day:02d}"
        else:
            final_date = None

        current_product_code = None
        current_product_name = None
        data = []

        for _, row in raw_df.iterrows():
            col_a = str(row[0]).strip() if pd.notna(row[0]) else ""
            col_b = str(row[1]).strip() if pd.notna(row[1]) else ""
            col_d = row[3] if pd.notna(row[3]) else None

            if "貨品編號" in col_a and "貨品名稱" in col_a:
                match = re.search(r'貨品編號[:：]([A-Z0-9\-]+)\s+貨品名稱[:：](.+)', col_a)
                if match:
                    current_product_code = match.group(1).strip()
                    current_product_name = match.group(2).strip()
                continue

            if "小計" in col_a or "小計" in col_b:
                continue

            if col_a.isdigit() and col_b and isinstance(col_d, (int, float)):
                data.append([
                    col_a, col_b, final_date,
                    current_product_code, current_product_name,
                    int(col_d)
                ])

        df_cleaned = pd.DataFrame(data, columns=[
            "Customer Code", "Customer Name", "Date",
            "Product Code", "Product Name", "Quantity"
        ])

        mapping_customer = pd.read_excel(mapping_file, sheet_name="Customer Mapping")
        mapping_customer = mapping_customer[[
            "ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Customer_No__c")

        df_cleaned = df_cleaned.merge(
            mapping_customer,
            left_on="Customer Code",
            right_on="ASI_CRM_Offtake_Customer_No__c",
            how="left"
        )

        df_cleaned["Customer Code"] = df_cleaned["ASI_CRM_JDE_Cust_No_Formula__c"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        df_cleaned.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"], inplace=True)

        mapping_sku = pd.read_excel(mapping_file, sheet_name="SKU Mapping")
        mapping_sku = mapping_sku[[
            "ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Product__c")

        df_cleaned = df_cleaned.merge(
            mapping_sku,
            left_on="Product Code",
            right_on="ASI_CRM_Offtake_Product__c",
            how="left"
        )

        product_code_index = df_cleaned.columns.get_loc("Product Code")
        df_cleaned.insert(product_code_index, "PRT Product Code", df_cleaned["ASI_CRM_SKU_Code__c"].astype(str).str.strip())
        df_cleaned.drop(columns=["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"], inplace=True)

        df_cleaned.insert(0, "Column1", "INV")
        df_cleaned.insert(1, "Column2", "U")
        df_cleaned.insert(2, "Column3", "30010010")
        df_cleaned.insert(3, "Column4", "酒倉 ON")

        st.write("✅ Processed Data Preview:")
        st.dataframe(df_cleaned)

        output_filename = "30010010 transformation.xlsx"
        df_cleaned.to_excel(output_filename, index=False, header=False)

        with open(output_filename, "rb") as f:
            st.download_button(label="📥 Download Processed File", data=f, file_name=output_filename)

elif transformation_choice == "30010013 酒田":
    raw_data_file = st.file_uploader("Upload Raw Sales Data", type=["xlsx", "xls"], key="sakata_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="sakata_mapping")

    if raw_data_file and mapping_file:
        raw_df = pd.read_excel(raw_data_file, sheet_name=0, header=None)  # Use first sheet

        # Extract ROC date from cell A5
        date_string = str(raw_df.iloc[4, 0])
        match = re.search(r'至\s*(\d{3}/\d{2}/\d{2})', date_string)
        if match:
            roc_date = match.group(1)
            year, month, day = map(int, roc_date.split('/'))
            final_date = f"{year + 1911}{month:02d}{day:02d}"
        else:
            final_date = None

        current_product_code = None
        current_product_name = None
        data = []

        for _, row in raw_df.iterrows():
            col_a = str(row[0]).strip() if pd.notna(row[0]) else ""
            col_b = str(row[1]).strip() if pd.notna(row[1]) else ""
            col_f = row.iloc[5] if len(row) > 5 and pd.notna(row.iloc[5]) else None  # SAFE

            if "貨品編號" in col_a and "貨品名稱" in col_a:
                match = re.search(r'貨品編號[:：]([A-Z0-9\-]+)\s+貨品名稱[:：](.+)', col_a)
                if match:
                    current_product_code = match.group(1).strip()
                    current_product_name = match.group(2).strip()
                continue

            if "小計" in col_a or "小計" in col_b:
                continue

            if col_a and col_a.startswith(("D", "C", "E", "M", "P")):
                if col_f and isinstance(col_f, (int, float)):
                    data.append([
                        col_a, col_b, final_date,
                        current_product_code, current_product_name,
                        int(col_f)
                    ])

        df_cleaned = pd.DataFrame(data, columns=[
            "Customer Code", "Customer Name", "Date",
            "Product Code", "Product Name", "Quantity"
        ])

        # Load customer mapping
        mapping_customer = pd.read_excel(mapping_file, sheet_name="Customer Mapping")
        mapping_customer = mapping_customer[[
            "ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Customer_No__c")

        df_cleaned = df_cleaned.merge(
            mapping_customer,
            left_on="Customer Code",
            right_on="ASI_CRM_Offtake_Customer_No__c",
            how="left"
        )

        df_cleaned["Customer Code"] = df_cleaned["ASI_CRM_JDE_Cust_No_Formula__c"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        df_cleaned.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"], inplace=True)

        # Load SKU mapping
        mapping_sku = pd.read_excel(mapping_file, sheet_name="SKU Mapping")
        mapping_sku = mapping_sku[[
            "ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Product__c")

        df_cleaned = df_cleaned.merge(
            mapping_sku,
            left_on="Product Code",
            right_on="ASI_CRM_Offtake_Product__c",
            how="left"
        )

        product_code_index = df_cleaned.columns.get_loc("Product Code")
        df_cleaned.insert(product_code_index, "PRT Product Code", df_cleaned["ASI_CRM_SKU_Code__c"].astype(str).str.strip())
        df_cleaned.drop(columns=["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"], inplace=True)

        # Insert fixed identifier columns
        df_cleaned.insert(0, "Column1", "INV")
        df_cleaned.insert(1, "Column2", "U")
        df_cleaned.insert(2, "Column3", "30010013")
        df_cleaned.insert(3, "Column4", "酒田 ON")

        st.write("✅ Processed Data Preview:")
        st.dataframe(df_cleaned)

        output_filename = "30010013 transformation.xlsx"
        df_cleaned.to_excel(output_filename, index=False, header=False)

        with open(output_filename, "rb") as f:
            st.download_button(label="📥 Download Processed File", data=f, file_name=output_filename)

elif transformation_choice == "30010059 誠邦有限公司":
    raw_data_file = st.file_uploader("Upload Raw Sales Data", type=["xlsx"], key="raw_30010059")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="mapping_30010059")

    if raw_data_file is not None and mapping_file is not None:
        raw_df = pd.read_excel(raw_data_file, sheet_name=0, header=None)

        data = []
        current_product_code = None
        current_product_name = None

        for _, row in raw_df.iterrows():
            col_a = str(row[0]).strip() if pd.notna(row[0]) else ""
            col_b = str(row[1]).strip() if pd.notna(row[1]) else ""
            col_c = str(row[2]).strip() if pd.notna(row[2]) else ""
            col_d = str(row[3]).strip() if pd.notna(row[3]) else ""
            col_e = row[4] if pd.notna(row[4]) else None

            # Clean invisible characters in col_a
            col_a_clean = col_a.replace('\u3000', ' ').replace('\xa0', ' ').strip()

            if "貨品編號:" in col_a_clean:
                match = re.search(r"貨品編號:\s*\[([^\]]+)\]\s*(.+)", col_a_clean)
                if match:
                    current_product_code = match.group(1).strip()
                    current_product_name = match.group(2).strip()
                continue

            if "合計" in col_a_clean or "小計" in col_a_clean:
                continue

            if re.match(r"\d{4}/\d{2}/\d{2}", col_a_clean) and col_c and isinstance(col_e, (int, float)):
                try:
                    y, m, d = map(int, col_a_clean.split("/"))
                    gregorian_date = f"{y + 1911}{m:02d}{d:02d}"
                except:
                    gregorian_date = col_a_clean

                data.append([
                    col_c, col_d, gregorian_date,
                    current_product_code, current_product_name,
                    int(col_e)
                ])

        df_cleaned = pd.DataFrame(data, columns=[
            "Customer Code", "Customer Name", "Date",
            "Product Code", "Product Name", "Quantity"
        ])

        dfs_mapping = {
            sheet: pd.read_excel(mapping_file, sheet_name=sheet)
            for sheet in pd.ExcelFile(mapping_file).sheet_names
        }

        df_customer_mapping = dfs_mapping["Customer Mapping"]
        df_customer_mapping = df_customer_mapping[[
            "ASI_CRM_Offtake_Customer_No__c",
            "ASI_CRM_JDE_Cust_No_Formula__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Customer_No__c")

        df_cleaned = df_cleaned.merge(
            df_customer_mapping,
            left_on="Customer Code",
            right_on="ASI_CRM_Offtake_Customer_No__c",
            how="left"
        )
        df_cleaned["Customer Code"] = df_cleaned["ASI_CRM_JDE_Cust_No_Formula__c"].astype(str).str.replace(r"\.0$", "", regex=True)
        df_cleaned.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"], inplace=True)

        df_sku_mapping = dfs_mapping["SKU Mapping"]
        df_sku_mapping = df_sku_mapping[[
            "ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Product__c")

        df_cleaned = df_cleaned.merge(
            df_sku_mapping,
            left_on="Product Code",
            right_on="ASI_CRM_Offtake_Product__c",
            how="left"
        )
        product_index = df_cleaned.columns.get_loc("Product Code")
        df_cleaned.insert(product_index, "PRT Product Code", df_cleaned["ASI_CRM_SKU_Code__c"].astype(str).str.strip())
        df_cleaned.drop(columns=["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"], inplace=True)

        fixed_df = pd.DataFrame({
            "Column1": ["INV"] * len(df_cleaned),
            "Column2": ["U"] * len(df_cleaned),
            "Column3": ["30010059"] * len(df_cleaned),
            "Column4": ["誠邦有限公司"] * len(df_cleaned)
        })

        df_final = pd.concat([fixed_df, df_cleaned], axis=1)

        st.write("✅ Processed Data Preview:")
        st.dataframe(df_final)

        output_filename = "processed_30010059.xlsx"
        df_final.to_excel(output_filename, index=False, header=False)

        with open(output_filename, "rb") as f:
            st.download_button(label="📥 Download Processed File", data=f, file_name=output_filename)

elif transformation_choice == "30010315 圳程":
    uploaded_file = st.file_uploader("Upload Raw Sales Data", type=["xlsx"], key="zhengcheng_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="zhengcheng_mapping")

    if uploaded_file and mapping_file:
        df_raw = pd.read_excel(uploaded_file, header=None)

        data = []
        current_product_code = None
        current_product_name = None

        # Extract date from cell A5 using the last ~ date
        raw_date_cell = str(df_raw.iloc[4, 0])
        match = re.findall(r'~\s*(\d{3}/\d{2}/\d{2})', raw_date_cell)
        if match:
            last_date = match[-1]
            y, m, d = map(int, last_date.split('/'))
            converted_date = f"{y + 1911}{m:02d}{d:02d}"
        else:
            converted_date = None

        for _, row in df_raw.iterrows():
            col_a = str(row[0]).strip() if pd.notna(row[0]) else ""
            col_b = str(row[1]).strip() if pd.notna(row[1]) else ""
            col_c = str(row[2]).strip() if pd.notna(row[2]) else ""
            col_d = row[3] if pd.notna(row[3]) else None

            # Match product code and name
            if "貨品編號" in col_a and "貨品名稱" in col_a:
                match = re.search(r"貨品編號[:：]([A-Z0-9\-]+)\s+貨品名稱[:：](.+)", col_a)
                if match:
                    current_product_code = match.group(1).strip()
                    current_product_name = match.group(2).strip()
                continue

            if "小計" in col_a or "小計" in col_b:
                continue

            if col_a and col_b and isinstance(col_d, (int, float)):
                data.append([
                    col_a, col_b, converted_date,
                    current_product_code, current_product_name,
                    int(col_d)
                ])

        df = pd.DataFrame(data, columns=[
            "Customer Code", "Customer Name", "Date",
            "Product Code", "Product Name", "Quantity"
        ])

        # Load mappings
        dfs_mapping = {
            sheet: pd.read_excel(mapping_file, sheet_name=sheet)
            for sheet in pd.ExcelFile(mapping_file).sheet_names
        }

        df_customer = dfs_mapping["Customer Mapping"]
        df_customer = df_customer[df_customer["ASI_CRM_Mapping_Cust_No__c"] == 30010315]
        df_customer = df_customer[["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"]].drop_duplicates()

        df = df.merge(df_customer, left_on="Customer Code", right_on="ASI_CRM_Offtake_Customer_No__c", how="left")
        df["Customer Code"] = df["ASI_CRM_JDE_Cust_No_Formula__c"].astype(str).str.replace(r"\.0$", "", regex=True)
        df.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"], inplace=True)

        df_sku = dfs_mapping["SKU Mapping"]
        df_sku = df_sku[["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"]].drop_duplicates()

        df = df.merge(df_sku, left_on="Product Code", right_on="ASI_CRM_Offtake_Product__c", how="left")
        insert_index = df.columns.get_loc("Product Code")
        df.insert(insert_index, "PRT Product Code", df["ASI_CRM_SKU_Code__c"].astype(str).str.strip())
        df.drop(columns=["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"], inplace=True)

        df.insert(0, "Col_1", "INV")
        df.insert(1, "Col_2", "U")
        df.insert(2, "Col_3", "30010315")
        df.insert(3, "Col_4", "圳程有限公司")

        st.write("✅ Processed Data Preview:")
        st.dataframe(df)

        output_filename = "30010315_transformation.xlsx"
        df.to_excel(output_filename, index=False, header=False)

        with open(output_filename, "rb") as f:
            st.download_button(label="📥 Download Processed File", data=f, file_name=output_filename)

