import streamlit as st
import pandas as pd
import re

# Streamlit app title
st.title("📊 WS Transformation")
st.write("Upload an Excel file and choose the transformation format.")

# Select transformation format
transformation_choice = st.selectbox("Select Transformation Format:", ["30010085 宏酒樽 (夜)", "30010203 宏酒樽 (日)", "30010061 向日葵", "30010010 酒倉盛豐行", "30010013 酒田", "30010059 誠邦有限公司", "30010315 圳程", "30030088 九久", "30020145 鏵錡", "30010199 振泰 OFF", "30010176 振泰 ON", "30030094 和易 ON", "33001422 和易 OFF"])

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
            output_filename = "30010085 transformation.xlsx"
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

            # Clean and normalize SKU columns
            df_transformed["Product Code"] = df_transformed["Product Code"].astype(str).str.strip().str.upper()
            df_sku_mapping["ASI_CRM_Offtake_Product__c"] = df_sku_mapping["ASI_CRM_Offtake_Product__c"].astype(str).str.strip().str.upper()

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
            output_filename = "30010203 transformation.xlsx"
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

            if isinstance(row[0], str) and '客戶名稱' in row[0]:
                cleaned_text = re.sub(r'[\u200b\ufeff]', '', row[0]).strip()
                match = re.search(r'客戶編號[:：]\s*([\d\-]+).*客戶名稱[:：]\s*(.*)', cleaned_text)
                if match:
                    current_customer_code = match.group(1).strip()
                    current_customer = match.group(2).strip()

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

        # --- ✅ CUSTOMER MAPPING ---
        dfs_mapping = {
            sheet: pd.read_excel(mapping_file, sheet_name=sheet)
            for sheet in pd.ExcelFile(mapping_file).sheet_names
        }

        df_customer = dfs_mapping["Customer Mapping"]
        df_customer = df_customer[[
            "ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Customer_No__c")

        result_df = result_df.merge(
            df_customer,
            left_on="Customer Code",
            right_on="ASI_CRM_Offtake_Customer_No__c",
            how="left"
        )

        result_df["Customer Code"] = result_df["ASI_CRM_JDE_Cust_No_Formula__c"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        result_df.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"], inplace=True)

        # --- ✅ SKU MAPPING ---
        df_sku = dfs_mapping["SKU Mapping"]
        df_sku = df_sku[[
            "ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Product__c")

        result_df = result_df.merge(
            df_sku,
            left_on="Product Code",
            right_on="ASI_CRM_Offtake_Product__c",
            how="left"
        )

        product_index = result_df.columns.get_loc("Product Code")
        result_df.insert(product_index, "PRT Product Code", result_df["ASI_CRM_SKU_Code__c"].astype(str).str.strip())
        result_df.drop(columns=["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"], inplace=True)

        # Preview data in Streamlit
        st.write("✅ Processed Data Preview:")
        st.dataframe(result_df)

        output_filename = "30010061 transformation.xlsx"
        result_df.to_excel(output_filename, index=False, header=False)

        with open(output_filename, "rb") as f:
            st.download_button(label="📥 Download Processed File", data=f, file_name=output_filename)

elif transformation_choice == "30010010 酒倉盛豐行":
    raw_data_file = st.file_uploader("Upload Raw Sales Data", type=["xlsx"], key="sakakura_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="sakakura_mapping")

    if raw_data_file and mapping_file:
        raw_df = pd.read_excel(raw_data_file, sheet_name=0, header=None)
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

            if col_a and col_b and isinstance(col_d, (int, float)) and current_product_code:
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

            if re.match(r'^[A-Z]', col_a):  # Allow any valid Latin-starting code
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
        import re
        import pandas as pd

        raw_df = pd.read_excel(raw_data_file, sheet_name=0, header=None)

        # Step 1: Detect format A or B (based on column B content for first date match)
        offset = 0
        for i in range(10, len(raw_df)):
            row = raw_df.iloc[i]
            col_a = str(row[0]).strip() if pd.notna(row[0]) else ""
            col_b = str(row[1]).strip() if pd.notna(row[1]) else ""
            if re.match(r"\d{4}/\d{2}/\d{2}|\d{3}/\d{2}/\d{2}", col_a):
                if col_b.startswith("\u92b7"):  # 銷
                    offset = 0  # Format A
                else:
                    offset = 1  # Format B
                break

        # Step 2: Extract product transactions using appropriate offset
        data = []
        current_product_code = None
        current_product_name = None
        found_first_product = False

        for i in range(len(raw_df)):
            row = raw_df.iloc[i]
            col_a = str(row[0]).strip() if pd.notna(row[0]) else ""
            col_b = str(row[1 - offset]).strip() if pd.notna(row[1 - offset]) else ""
            col_c = str(row[2 - offset]).strip() if pd.notna(row[2 - offset]) else ""
            col_d = str(row[3 - offset]).strip() if pd.notna(row[3 - offset]) else ""
            col_e = row[4 - offset] if pd.notna(row[4 - offset]) else None

            col_a_clean = col_a.replace('\u3000', ' ').replace('\xa0', ' ').strip()

            # Match both 【】 and []
            if "貨品編號:" in col_a_clean:
                match = re.search(r"貨品編號:\s*[\[\【]([^\]\】]+)[\]\】]\s*(.+)", col_a_clean)
                if match:
                    current_product_code = match.group(1).strip()
                    current_product_name = match.group(2).strip()
                    found_first_product = True
                continue

            if not found_first_product:
                continue

            if "合計" in col_a_clean or "小計" in col_a_clean:
                continue

            if col_c and isinstance(col_e, (int, float)) and current_product_code and current_product_name:
                try:
                    y, m, d = map(int, col_a_clean.split("/"))
                    if y < 1911:
                        y += 1911
                    gregorian_date = f"{y}{m:02d}{d:02d}"
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

        # Load mapping file and merge
        dfs_mapping = {
            sheet: pd.read_excel(mapping_file, sheet_name=sheet)
            for sheet in pd.ExcelFile(mapping_file).sheet_names
        }

        # Customer mapping
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
        df_cleaned["Customer Code"] = df_cleaned["ASI_CRM_JDE_Cust_No_Formula__c"].astype(str).str.replace(r"\\.0$", "", regex=True)
        df_cleaned.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"], inplace=True)

        # SKU mapping
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

        # Add fixed columns
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
            st.download_button(label="📅 Download Processed File", data=f, file_name=output_filename)


elif transformation_choice == "30010315 圳程":
    raw_data_file = st.file_uploader("Upload Raw Sales Data", type=["xlsx"], key="zc_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="zc_mapping")

    if raw_data_file and mapping_file:
        import openpyxl

        wb = openpyxl.load_workbook(raw_data_file, data_only=True)
        ws = wb.active

        # Try B3, then B4 if B3 is empty
        report_date_raw = ""
        for cell in ["B3", "B4"]:
            val = ws[cell].value
            if val:
                report_date_raw = str(val).strip()
                break

        # Parse the date string if available
        report_date = ""
        if "~" in report_date_raw:
            right_date = report_date_raw.split("~")[-1].strip()
            if len(right_date.split("/")) == 3:
                y, m, d = right_date.split("/")
                report_date = f"{int(y):04}{int(m):02}{int(d):02}"



        records = []
        product_name = product_code = customer_name = customer_code = None

        for i in range(ws.max_row):
            b = str(ws.cell(row=i+1, column=2).value).strip() if ws.cell(row=i+1, column=2).value else ""
            c = str(ws.cell(row=i+1, column=3).value).strip() if ws.cell(row=i+1, column=3).value else ""
            e = ws.cell(row=i+1, column=5).value if ws.cell(row=i+1, column=5).value else None

            if "(" in b and ")" in b:
                last_open = b.rfind("(")
                last_close = b.rfind(")")
                code = b[last_open + 1 : last_close]
                name = b[:last_open].strip()

                if i+2 < ws.max_row and str(ws.cell(row=i+2, column=2).value).strip() == "單據類別":
                    customer_name = name
                    customer_code = code
                else:
                    product_name = name
                    product_code = code

            if b == "出貨單" and c and isinstance(e, (int, float)):
                records.append({
                    "Customer Code": customer_code,
                    "Customer Name": customer_name,
                    "Date": report_date,
                    "Product Code": product_code,
                    "Product Name": product_name,
                    "Quantity": int(e),
                    "Document Number": c
                })

        df_transformed = pd.DataFrame(records)
        df_transformed.insert(0, "Column1", "INV")
        df_transformed.insert(1, "Column2", "U")
        df_transformed.insert(2, "Column3", "30010315")
        df_transformed.insert(3, "Column4", "圳程有限公司")

        # Load mappings
        dfs_mapping = {
            sheet: pd.read_excel(mapping_file, sheet_name=sheet)
            for sheet in pd.ExcelFile(mapping_file).sheet_names
        }

        # Customer mapping
        df_customer = dfs_mapping["Customer Mapping"]
        df_customer = df_customer[[
            "ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Customer_No__c")

        df_transformed = df_transformed.merge(
            df_customer,
            left_on="Customer Code",
            right_on="ASI_CRM_Offtake_Customer_No__c",
            how="left"
        )

        df_transformed["Customer Code"] = df_transformed["ASI_CRM_JDE_Cust_No_Formula__c"].astype(str).str.replace(r"\.0$", "", regex=True)
        df_transformed.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"], inplace=True)

        # SKU mapping
        df_sku = dfs_mapping["SKU Mapping"]
        df_sku = df_sku[[
            "ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Product__c")

        df_transformed = df_transformed.merge(
            df_sku,
            left_on="Product Code",
            right_on="ASI_CRM_Offtake_Product__c",
            how="left"
        )

        product_index = df_transformed.columns.get_loc("Product Code")
        df_transformed.insert(product_index, "PRT Product Code", df_transformed["ASI_CRM_SKU_Code__c"].astype(str).str.strip())
        df_transformed.drop(columns=["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"], inplace=True)

        # Reorder for consistency
        column_order = ["Column1", "Column2", "Column3", "Column4", "Customer Code", "Customer Name", "Date", "PRT Product Code", "Product Code", "Product Name", "Quantity", "Document Number"]
        df_transformed = df_transformed[column_order]

        st.write("✅ Processed Data Preview:")
        st.dataframe(df_transformed)

        output_filename = "30010315_transformation.xlsx"
        df_transformed.to_excel(output_filename, index=False, header=False)

        with open(output_filename, "rb") as f:
            st.download_button(label="📥 Download Processed File", data=f, file_name=output_filename)
            
elif transformation_choice == "30030088 九久":
    raw_data_file = st.file_uploader("Upload Raw Sales Data", type=["xlsx"], key="jj_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="jj_mapping")

    if raw_data_file and mapping_file:
        import openpyxl

        df_raw = pd.read_excel(raw_data_file, sheet_name=0, header=None)
        extracted_data = []

        i = 0
        while i < len(df_raw):
            row = df_raw.iloc[i, 0]

            if isinstance(row, str) and row.startswith("貨品編號:"):
                product_code = row.replace("貨品編號:", "").split()[0].strip()
                product_name = row.replace("貨品編號:", "").split(maxsplit=1)[1].strip() if len(row.split()) > 1 else ""

                data_start = i + 5
                while data_start < len(df_raw):
                    entry = df_raw.iloc[data_start]

                    if isinstance(entry[0], str) and entry[0].startswith("貨品編號:"):
                        break

                    # ✅ Skip if inbound: check if column E is '進貨單'
                    if str(entry[4]).strip() == "進貨單":
                        data_start += 1
                        continue

                    # Check if entry is valid (i.e., not empty)
                    if pd.isna(entry[0]) or pd.isna(entry[1]) or pd.isna(entry[2]):
                        data_start += 1
                        continue

                    # Initialize return flag
                    is_return = False

                    # ✅ Skip if inbound: check if column E is '進貨單'
                    # ✅ If it's a return '銷退單', we mark it and negate quantity later
                    doc_type = str(entry[4]).strip()
                    if doc_type == "進貨單":
                        data_start += 1
                        continue
                    elif doc_type == "銷退單":
                        is_return = True

                    try:
                        report_date = entry[0]
                        document_number = entry[1]
                        customer_code = entry[2]
                        customer_name = entry[3]
                        quantity = entry[6]

                        if pd.notna(quantity) and isinstance(quantity, (int, float)):
                            if is_return:
                                quantity = -abs(int(quantity))  # Ensure it's negative
                            extracted_data.append({
                                "Customer Code": str(customer_code).strip().split(".")[0],
                                "Customer Name": str(customer_name).strip(),
                                "Date": report_date,
                                "Product Code": product_code,
                                "Product Name": product_name,
                                "Quantity": int(quantity),
                                "Document Number": document_number
                            })
                    except Exception:
                        pass
                    data_start += 1
            i += 1


        df_transformed = pd.DataFrame(extracted_data)

        # Convert Minguo date to Gregorian YYYYMMDD
        def convert_minguo_date(minguo_str):
            try:
                parts = str(minguo_str).split('/')
                if len(parts) != 3:
                    return None
                year = int(parts[0]) + 1911
                month = int(parts[1])
                day = int(parts[2])
                return f"{year:04d}{month:02d}{day:02d}"
            except Exception:
                return None

        df_transformed["Date"] = df_transformed["Date"].apply(convert_minguo_date)

        # Add fixed columns
        df_transformed.insert(0, "Column4", "九久")
        df_transformed.insert(0, "Column3", "30030088")
        df_transformed.insert(0, "Column2", "U")
        df_transformed.insert(0, "Column1", "INV")

        # Load mapping sheets
        dfs_mapping = {
            sheet: pd.read_excel(mapping_file, sheet_name=sheet)
            for sheet in pd.ExcelFile(mapping_file).sheet_names
        }

        # Customer mapping
        df_customer = dfs_mapping["Customer Mapping"]
        df_customer = df_customer[[
            "ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Customer_No__c")

        df_transformed = df_transformed.merge(
            df_customer,
            left_on="Customer Code",
            right_on="ASI_CRM_Offtake_Customer_No__c",
            how="left"
        )

        df_transformed["Customer Code"] = df_transformed["ASI_CRM_JDE_Cust_No_Formula__c"].astype(str).str.replace(r"\.0$", "", regex=True)
        df_transformed.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"], inplace=True)

        # SKU mapping
        df_sku = dfs_mapping["SKU Mapping"]
        df_sku = df_sku[[
            "ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Product__c")

        df_transformed = df_transformed.merge(
            df_sku,
            left_on="Product Code",
            right_on="ASI_CRM_Offtake_Product__c",
            how="left"
        )

        product_index = df_transformed.columns.get_loc("Product Code")
        df_transformed.insert(product_index, "PRT Product Code", df_transformed["ASI_CRM_SKU_Code__c"].astype(str).str.strip())
        df_transformed.drop(columns=["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"], inplace=True)

        # Final column order
        column_order = ["Column1", "Column2", "Column3", "Column4", "Customer Code", "Customer Name", "Date", "PRT Product Code", "Product Code", "Product Name", "Quantity", "Document Number"]
        df_transformed = df_transformed[column_order]

        st.write("✅ Processed Data Preview:")
        st.dataframe(df_transformed)

        output_filename = "30030088_transformation.xlsx"
        df_transformed.to_excel(output_filename, index=False, header=False)

        with open(output_filename, "rb") as f:
            st.download_button(label="📥 Download Processed File", data=f, file_name=output_filename)


elif transformation_choice == "30020145 鏵錡":
    raw_data_file = st.file_uploader("Upload Raw Sales Data", type=["xlsx"], key="30020145_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="30020145_mapping")

    if raw_data_file and mapping_file:
        import pandas as pd
        import re

        def extract_product_data_from_workbook(file):
            xls = pd.ExcelFile(file)
            combined_data = []

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                merged_cell_value = str(df.iloc[2, 0])
                product_match = re.search(r"貨品編號[:：]([A-Z0-9\-]+)\s+(.*)", merged_cell_value)

                if not product_match:
                    continue

                product_code = product_match.group(1).strip()
                product_name = product_match.group(2).strip()

                df_data = df.iloc[8:, :8].copy()
                df_data.columns = ["Date", "Document No", "Customer Code", "Distributor", "Customer Name", "Quantity", "Unit", "Note"]

                for _, row in df_data.iterrows():
                    if pd.isna(row["Date"]) or pd.isna(row["Customer Code"]) or pd.isna(row["Quantity"]):
                        continue

                    combined_data.append({
                        "Customer Code": row["Customer Code"],
                        "Customer Name": row["Customer Name"],
                        "Date": row["Date"],
                        "Product Code": product_code,
                        "Product Name": product_name,
                        "Quantity": row["Quantity"],
                        "Document No": row["Document No"]
                    })

            return pd.DataFrame(combined_data)

        def convert_minguo_to_gregorian(date_str):
            try:
                parts = str(date_str).split("/")
                if len(parts) != 3:
                    return None
                year = int(parts[0]) + 1911
                month = int(parts[1])
                day = int(parts[2])
                return f"{year:04d}{month:02d}{day:02d}"
            except:
                return None

        df_combined = extract_product_data_from_workbook(raw_data_file)
        df_combined["Date"] = df_combined["Date"].apply(convert_minguo_to_gregorian)

        # Load mapping sheets
        dfs_mapping = {
            sheet: pd.read_excel(mapping_file, sheet_name=sheet)
            for sheet in pd.ExcelFile(mapping_file).sheet_names
        }

        # Customer Mapping
        df_customer = dfs_mapping["Customer Mapping"]
        df_customer = df_customer[[
            "ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Customer_No__c")

        df_combined = df_combined.merge(
            df_customer,
            left_on="Customer Code",
            right_on="ASI_CRM_Offtake_Customer_No__c",
            how="left"
        )

        df_combined["Customer Code"] = df_combined["ASI_CRM_JDE_Cust_No_Formula__c"].astype(str).str.replace(r"\.0$", "", regex=True)
        df_combined.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"], inplace=True)

        # SKU Mapping
        df_sku = dfs_mapping["SKU Mapping"]
        df_sku = df_sku[[
            "ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Product__c")

        df_combined = df_combined.merge(
            df_sku,
            left_on="Product Code",
            right_on="ASI_CRM_Offtake_Product__c",
            how="left"
        )

        product_index = df_combined.columns.get_loc("Product Code")
        df_combined.insert(product_index, "PRT Product Code", df_combined["ASI_CRM_SKU_Code__c"].astype(str).str.strip())
        df_combined.drop(columns=["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"], inplace=True)

        # Insert fixed columns
        df_combined.insert(0, "Column4", "任我行")
        df_combined.insert(0, "Column3", "30020145")
        df_combined.insert(0, "Column2", "U")
        df_combined.insert(0, "Column1", "INV")

        # Preview result
        st.write("✅ Processed Data Preview:")
        st.dataframe(df_combined)

        output_filename = "30020145_transformation.xlsx"
        df_combined.to_excel(output_filename, index=False, header=False)

        with open(output_filename, "rb") as f:
            st.download_button(label="📥 Download Processed File", data=f, file_name=output_filename)

elif transformation_choice == "30010199 振泰 OFF":
    raw_data_file = st.file_uploader("Upload Raw Sales Data", type=["xls"], key="zhen_tai_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="zhen_tai_mapping")

    if raw_data_file is not None and mapping_file is not None:
        def extract_from_date_sheets(file):
            xls = pd.ExcelFile(file)
            all_data = []
            sheet_dates = {}

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet_name, header=None)
                product_code = None
                product_name = None

                # ✅ Skip sheet if A5 is missing
                if df.shape[0] <= 4 or pd.isna(df.iloc[4, 0]):
                    continue

                # Extract date from A5
                raw_date_cell = str(df.iloc[4, 0])

                if "至" in raw_date_cell:
                    raw_date = raw_date_cell.split("至")[1].strip()
                    try:
                        parts = raw_date.split("/")
                        year = int(parts[0]) + 1911
                        month = int(parts[1])
                        day = int(parts[2])
                        formatted_date = f"{year:04d}{month:02d}{day:02d}"
                    except:
                        formatted_date = None
                else:
                    formatted_date = None
                sheet_dates[sheet_name] = formatted_date

                for i in range(len(df)):
                    cell_value = str(df.iloc[i, 0]).strip()
                    if cell_value.startswith("貨品編號:"):
                        rest = cell_value.replace("貨品編號:", "", 1).strip()
                        parts = rest.split("貨品名稱:")
                        product_code = parts[0].strip()
                        product_name = parts[1].strip() if len(parts) > 1 else ""
                        continue
                    if "小計" in cell_value or product_code is None:
                        continue

                    customer_code = str(df.iloc[i, 0]).strip()
                    customer_name = str(df.iloc[i, 1]).strip() if pd.notna(df.iloc[i, 1]) else ""
                    quantity = df.iloc[i, 2] if pd.notna(df.iloc[i, 2]) else None

                    if customer_code and quantity and isinstance(quantity, (int, float)):
                        all_data.append({
                            "Sheet": sheet_name,
                            "Customer Code": customer_code,
                            "Customer Name": customer_name,
                            "Date": formatted_date,
                            "Product Code": product_code,
                            "Product Name": product_name,
                            "Quantity": quantity
                        })

            return pd.DataFrame(all_data)

        df = extract_from_date_sheets(raw_data_file)

        # Mapping setup
        dfs_mapping = {
            sheet: pd.read_excel(mapping_file, sheet_name=sheet)
            for sheet in pd.ExcelFile(mapping_file).sheet_names
        }

        # Filter customer mapping
        df_customer_mapping = dfs_mapping["Customer Mapping"]
        df_customer_mapping = df_customer_mapping[
            df_customer_mapping["ASI_CRM_Mapping_Cust_No__c"] == 30010199
        ][["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"]].drop_duplicates()

        df = df.merge(
            df_customer_mapping,
            left_on="Customer Code",
            right_on="ASI_CRM_Offtake_Customer_No__c",
            how="left"
        )

        df["Customer Code"] = df["ASI_CRM_JDE_Cust_No_Formula__c"].astype(str).str.replace(r"\.0$", "", regex=True)
        df.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"], inplace=True)

        df_sku_mapping = dfs_mapping["SKU Mapping"]
        df_sku_mapping = df_sku_mapping[[
            "ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"
        ]].drop_duplicates()

        df = df.merge(
            df_sku_mapping,
            left_on="Product Code",
            right_on="ASI_CRM_Offtake_Product__c",
            how="left"
        )
        df.insert(df.columns.get_loc("Product Code"), "PRT Product Code", df["ASI_CRM_SKU_Code__c"].astype(str).str.strip())
        df.drop(columns=["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"], inplace=True)

        # Add 4 fixed columns
        df.insert(1, "Col1", "INV")
        df.insert(2, "Col2", "U")
        df.insert(3, "Col3", "30010199")
        df.insert(4, "Col4", "振泰 OFF")

        # Optional: Toggle by Month (📅 grouped by available months)
        available_months = sorted(set([d[:6] for d in df["Date"].dropna().astype(str)]))
        month_filter = st.radio("📅 Filter by Month:", ["All"] + available_months)

        if month_filter != "All":
            df = df[df["Date"].astype(str).str.startswith(month_filter)]

        # Final column order
        df = df[[
            "Sheet", "Col1", "Col2", "Col3", "Col4",
            "Customer Code", "Customer Name", "Date",
            "PRT Product Code", "Product Code", "Product Name", "Quantity"
        ]]

        st.write("✅ Processed Data Preview:")
        st.dataframe(df)

        st.download_button(
            label="📥 Download Processed File",
            data=df.to_csv(index=False),
            file_name="zhen_tai_processed.csv",
            mime="text/csv"
        )

elif transformation_choice == "30010176 振泰 ON":
    raw_data_file = st.file_uploader("Upload Raw Sales Data", type=["xls"], key="zhen_tai_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="zhen_tai_mapping")

    if raw_data_file is not None and mapping_file is not None:
        def extract_from_date_sheets(file):
            xls = pd.ExcelFile(file)
            all_data = []
            sheet_dates = {}

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet_name, header=None)
                product_code = None
                product_name = None

                # ✅ Skip sheet if A5 is missing
                if df.shape[0] <= 4 or pd.isna(df.iloc[4, 0]):
                    continue

                # Extract date from A5
                raw_date_cell = str(df.iloc[4, 0])

                if "至" in raw_date_cell:
                    raw_date = raw_date_cell.split("至")[1].strip()
                    try:
                        parts = raw_date.split("/")
                        year = int(parts[0]) + 1911
                        month = int(parts[1])
                        day = int(parts[2])
                        formatted_date = f"{year:04d}{month:02d}{day:02d}"
                    except:
                        formatted_date = None
                else:
                    formatted_date = None
                sheet_dates[sheet_name] = formatted_date

                for i in range(len(df)):
                    cell_value = str(df.iloc[i, 0]).strip()
                    if cell_value.startswith("貨品編號:"):
                        rest = cell_value.replace("貨品編號:", "", 1).strip()
                        parts = rest.split("貨品名稱:")
                        product_code = parts[0].strip()
                        product_name = parts[1].strip() if len(parts) > 1 else ""
                        continue
                    if "小計" in cell_value or product_code is None:
                        continue

                    customer_code = str(df.iloc[i, 0]).strip()
                    customer_name = str(df.iloc[i, 1]).strip() if pd.notna(df.iloc[i, 1]) else ""
                    quantity = df.iloc[i, 2] if pd.notna(df.iloc[i, 2]) else None

                    if customer_code and quantity and isinstance(quantity, (int, float)):
                        all_data.append({
                            "Sheet": sheet_name,
                            "Customer Code": customer_code,
                            "Customer Name": customer_name,
                            "Date": formatted_date,
                            "Product Code": product_code,
                            "Product Name": product_name,
                            "Quantity": quantity
                        })

            return pd.DataFrame(all_data)

        df = extract_from_date_sheets(raw_data_file)

        # Mapping setup
        dfs_mapping = {
            sheet: pd.read_excel(mapping_file, sheet_name=sheet)
            for sheet in pd.ExcelFile(mapping_file).sheet_names
        }

        # Filter customer mapping
        df_customer_mapping = dfs_mapping["Customer Mapping"]
        df_customer_mapping = df_customer_mapping[
            df_customer_mapping["ASI_CRM_Mapping_Cust_No__c"] == 30010199
        ][["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"]].drop_duplicates()

        df = df.merge(
            df_customer_mapping,
            left_on="Customer Code",
            right_on="ASI_CRM_Offtake_Customer_No__c",
            how="left"
        )

        df["Customer Code"] = df["ASI_CRM_JDE_Cust_No_Formula__c"].astype(str).str.replace(r"\.0$", "", regex=True)
        df.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"], inplace=True)

        df_sku_mapping = dfs_mapping["SKU Mapping"]
        df_sku_mapping = df_sku_mapping[[
            "ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"
        ]].drop_duplicates()

        df = df.merge(
            df_sku_mapping,
            left_on="Product Code",
            right_on="ASI_CRM_Offtake_Product__c",
            how="left"
        )
        df.insert(df.columns.get_loc("Product Code"), "PRT Product Code", df["ASI_CRM_SKU_Code__c"].astype(str).str.strip())
        df.drop(columns=["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"], inplace=True)

        # Add 4 fixed columns
        df.insert(1, "Col1", "INV")
        df.insert(2, "Col2", "U")
        df.insert(3, "Col3", "30010199")
        df.insert(4, "Col4", "振泰 OFF")

        # Optional: Toggle by Month (📅 grouped by available months)
        available_months = sorted(set([d[:6] for d in df["Date"].dropna().astype(str)]))
        month_filter = st.radio("📅 Filter by Month:", ["All"] + available_months)

        if month_filter != "All":
            df = df[df["Date"].astype(str).str.startswith(month_filter)]

        # Final column order
        df = df[[
            "Sheet", "Col1", "Col2", "Col3", "Col4",
            "Customer Code", "Customer Name", "Date",
            "PRT Product Code", "Product Code", "Product Name", "Quantity"
        ]]

        st.write("✅ Processed Data Preview:")
        st.dataframe(df)

        st.download_button(
            label="📥 Download Processed File",
            data=df.to_csv(index=False),
            file_name="zhen_tai_processed.csv",
            mime="text/csv"
        )

elif transformation_choice == "30030094 和易 ON":
    raw_data_file = st.file_uploader("Upload Raw Sales Data", type=["xls", "xlsx"], key="heyi_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xls", "xlsx"], key="heyi_mapping")

    if raw_data_file and mapping_file:
        raw_df = pd.read_excel(raw_data_file, sheet_name="Page 1", header=None)

        # Extract depletion rows with context
        extracted_data = []
        product_code = None
        product_name = None

        for idx, row in raw_df.iterrows():
            col0 = str(row[0]) if pd.notna(row[0]) else ""
            col3 = str(row[3]) if pd.notna(row[3]) else ""

            if col0.startswith("產品編號:"):
                product_code = col0.replace("產品編號:", "").strip()

            if col3.startswith("品名規格:"):
                product_name = col3.replace("品名規格:", "").strip()

            if str(row[3]).strip() == "銷貨（庫存）":
                report_date = row[0]
                document_number = row[1]
                customer_name = row[2]
                quantity = row[5]
                customer_code = row[9]

                if all(pd.notna([report_date, document_number, customer_name, quantity, customer_code])):
                    extracted_data.append({
                        "Customer Code": str(customer_code).strip(),
                        "Customer Name": str(customer_name).strip(),
                        "Date": report_date,
                        "Product Code": product_code,
                        "Product Name": product_name,
                        "Quantity": int(quantity),
                        "Document Number": document_number
                    })

        depletion_df = pd.DataFrame(extracted_data)

        # Add fixed columns
        depletion_df.insert(0, "INV", "INV")
        depletion_df.insert(1, "U", "U")
        depletion_df.insert(2, "Customer Group Code", "30030094")
        depletion_df.insert(3, "Customer Group Name", "和易 ON")

        # Mapping: Customer
        mapping_customer = pd.read_excel(mapping_file, sheet_name="Customer Mapping")
        mapping_customer = mapping_customer[[
            "ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Customer_No__c")

        depletion_df = depletion_df.merge(
            mapping_customer,
            left_on="Customer Code",
            right_on="ASI_CRM_Offtake_Customer_No__c",
            how="left"
        )

        depletion_df["Customer Code"] = depletion_df["ASI_CRM_JDE_Cust_No_Formula__c"].astype(str).str.replace(r"\\.0$", "", regex=True)
        depletion_df.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"], inplace=True)

        # Mapping: SKU
        mapping_sku = pd.read_excel(mapping_file, sheet_name="SKU Mapping")
        mapping_sku = mapping_sku[[
            "ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Product__c")

        depletion_df = depletion_df.merge(
            mapping_sku,
            left_on="Product Code",
            right_on="ASI_CRM_Offtake_Product__c",
            how="left"
        )

        product_index = depletion_df.columns.get_loc("Product Code")
        depletion_df.insert(product_index, "PRT Product Code", depletion_df["ASI_CRM_SKU_Code__c"].astype(str).str.strip())
        depletion_df.drop(columns=["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"], inplace=True)

        # Convert Minguo date to YYYYMMDD
        def convert_minguo_date(date_str):
            try:
                if isinstance(date_str, str) and '/' in date_str:
                    parts = date_str.strip().split('/')
                    year = int(parts[0]) + 1911
                    month = int(parts[1])
                    day = int(parts[2])
                    return f"{year:04d}{month:02d}{day:02d}"
                return date_str
            except:
                return date_str

        depletion_df["Date"] = depletion_df["Date"].apply(convert_minguo_date)

        st.write("✅ Processed Data Preview:")
        st.dataframe(depletion_df)

        output_filename = "30030094_transformation.xlsx"
        depletion_df.to_excel(output_filename, index=False, header=False)

        with open(output_filename, "rb") as f:
            st.download_button(label="📥 Download Processed File", data=f, file_name=output_filename)

elif transformation_choice == "33001422 和易 OFF":
    raw_data_file = st.file_uploader("Upload Raw Sales Data", type=["xls", "xlsx"], key="heyi_off_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xls", "xlsx"], key="heyi_off_mapping")

    if raw_data_file and mapping_file:
        raw_df = pd.read_excel(raw_data_file, sheet_name="Page 1", header=None)

        extracted_data = []
        product_code = None
        product_name = None

        for _, row in raw_df.iterrows():
            col0 = str(row[0]) if pd.notna(row[0]) else ""
            col3 = str(row[3]) if pd.notna(row[3]) else ""

            if col0.startswith("產品編號:"):
                product_code = col0.replace("產品編號:", "").strip()

            if col3.startswith("品名規格:"):
                product_name = col3.replace("品名規格:", "").strip()

            if str(row[3]).strip() in ["銷貨（庫存）", "銷貨退回"]:
                report_date = row[0]
                document_number = row[1]
                customer_name = row[2]
                quantity = row[5]
                customer_code = row[9]

                if all(pd.notna([report_date, document_number, customer_name, quantity, customer_code])):
                    qty = int(quantity)
                    if str(row[3]).strip() == "銷貨退回":
                        qty = -qty
                    extracted_data.append({
                        "Customer Code": str(customer_code).strip(),
                        "Customer Name": str(customer_name).strip(),
                        "Date": report_date,
                        "Product Code": product_code,
                        "Product Name": product_name,
                        "Quantity": qty,
                        "Document Number": document_number
                    })

        df_extracted = pd.DataFrame(extracted_data)

        # Add 4 fixed metadata columns
        df_extracted.insert(0, "INV", "INV")
        df_extracted.insert(1, "U", "U")
        df_extracted.insert(2, "Customer Group Code", "33001422")
        df_extracted.insert(3, "Customer Group Name", "和易 OFF")

        # Convert Minguo date to Gregorian
        def convert_minguo_date(date_str):
            try:
                if isinstance(date_str, str) and '/' in date_str:
                    year, month, day = map(int, date_str.split('/'))
                    return f"{year + 1911:04d}{month:02d}{day:02d}"
                return date_str
            except:
                return date_str

        df_extracted["Date"] = df_extracted["Date"].apply(convert_minguo_date)

        # Customer Mapping
        mapping_customer = pd.read_excel(mapping_file, sheet_name="Customer Mapping")
        mapping_customer = mapping_customer[[
            "ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Customer_No__c")

        df_extracted = df_extracted.merge(
            mapping_customer,
            left_on="Customer Code",
            right_on="ASI_CRM_Offtake_Customer_No__c",
            how="left"
        )

        df_extracted["Customer Code"] = df_extracted["ASI_CRM_JDE_Cust_No_Formula__c"].astype(str).str.replace(r"\.0$", "", regex=True)
        df_extracted.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"], inplace=True)

        # SKU Mapping
        mapping_sku = pd.read_excel(mapping_file, sheet_name="SKU Mapping")
        mapping_sku = mapping_sku[[
            "ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Product__c")

        df_extracted = df_extracted.merge(
            mapping_sku,
            left_on="Product Code",
            right_on="ASI_CRM_Offtake_Product__c",
            how="left"
        )

        product_index = df_extracted.columns.get_loc("Product Code")
        df_extracted.insert(product_index, "PRT Product Code", df_extracted["ASI_CRM_SKU_Code__c"].astype(str).str.strip())
        df_extracted.drop(columns=["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"], inplace=True)

        st.write("✅ Processed Data Preview:")
        st.dataframe(df_extracted)

        output_filename = "33001422_transformation.xlsx"
        df_extracted.to_excel(output_filename, index=False, header=False)

        with open(output_filename, "rb") as f:
            st.download_button(label="📥 Download Processed File", data=f, file_name=output_filename)

