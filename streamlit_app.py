import streamlit as st
import pandas as pd
import re

st.title("📊 WS Transformation")
st.write("Upload an Excel file and choose the transformation format.")

# Dropdown for transformation selection
transformation_choice = st.radio("Select Transformation Format:", ["30010010 酒倉盛豐行", "向日葵", "30010013 酒田"])

# ────────────────────────────── 30010010 酒倉盛豐行 ──────────────────────────────
if transformation_choice == "30010010 酒倉盛豐行":
    raw_file = st.file_uploader("Upload Raw Sales Data", type=["xlsx"], key="upload_001")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="mapping_001")

    if raw_file and mapping_file:
        df_raw = pd.read_excel(raw_file, sheet_name="給廠商", header=None)

        # Extract date
        date_text = str(df_raw.iloc[4, 0])
        match = re.search(r'至\s*(\d{3}/\d{2}/\d{2})', date_text)
        if match:
            year, month, day = map(int, match.group(1).split('/'))
            final_date = f"{year + 1911}{month:02d}{day:02d}"
        else:
            final_date = None

        current_product_code, current_product_name = None, None
        data = []

        for _, row in df_raw.iterrows():
            col_a = str(row[0]).strip() if pd.notna(row[0]) else ""
            col_b = str(row[1]).strip() if pd.notna(row[1]) else ""
            col_d = row[3] if pd.notna(row[3]) else None

            if "貨品編號" in col_a and "貨品名稱" in col_a:
                m = re.search(r'貨品編號[:：]([A-Z0-9\-]+)\s+貨品名稱[:：](.+)', col_a)
                if m:
                    current_product_code = m.group(1).strip()
                    current_product_name = m.group(2).strip()
                continue

            if "小計" in col_a or "小計" in col_b:
                continue

            if col_a.isdigit() and col_b and isinstance(col_d, (int, float)):
                data.append([col_a, col_b, final_date, current_product_code, current_product_name, int(col_d)])

        df = pd.DataFrame(data, columns=["Customer Code", "Customer Name", "Date", "Product Code", "Product Name", "Quantity"])

        # Customer mapping
        mapping = pd.read_excel(mapping_file, sheet_name="Customer Mapping")
        mapping = mapping[["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"]].drop_duplicates()
        df = df.merge(mapping, left_on="Customer Code", right_on="ASI_CRM_Offtake_Customer_No__c", how="left")
        df["Customer Code"] = df["ASI_CRM_JDE_Cust_No_Formula__c"].astype(str).str.replace(r"\.0$", "", regex=True)
        df.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"], inplace=True)

        # SKU Mapping
        sku_map = pd.read_excel(mapping_file, sheet_name="SKU Mapping")
        sku_map = sku_map[["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"]].drop_duplicates()
        df = df.merge(sku_map, left_on="Product Code", right_on="ASI_CRM_Offtake_Product__c", how="left")
        df.insert(df.columns.get_loc("Product Code"), "PRT Product Code", df["ASI_CRM_SKU_Code__c"].astype(str).str.strip())
        df.drop(columns=["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"], inplace=True)

        # Add fixed columns
        df.insert(0, "Column1", "INV")
        df.insert(1, "Column2", "U")
        df.insert(2, "Column3", "30010010")
        df.insert(3, "Column4", "酒倉 ON")

        st.write("✅ Preview:")
        st.dataframe(df)

        filename = "30010010 transformation.xlsx"
        df.to_excel(filename, index=False, header=False)
        with open(filename, "rb") as f:
            st.download_button("📥 Download File", f, file_name=filename)

# ────────────────────────────── 向日葵 ──────────────────────────────
elif transformation_choice == "向日葵":
    uploaded_file = st.file_uploader("Upload Raw Sales Data", type=["xlsx"], key="sunflower_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="sunflower_mapping")

    if uploaded_file is not None and mapping_file is not None:
        df = pd.read_excel(uploaded_file, header=None)
        data = []
        current_customer = current_customer_code = current_date = None

        for i in range(7, len(df)):
            row = df.iloc[i]
            if isinstance(row[0], str) and '客戶名稱' in row[0]:
                match = re.search(r'客戶編號[:：]\s*([\d\-]+).*客戶名稱[:：]\s*(.*)', row[0])
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

        df = pd.DataFrame(data, columns=['Customer Code', 'Customer Name', 'Date', 'Product Code', 'Product Name', 'Quantity'])
        df.insert(0, 'Column1', 'INV')
        df.insert(1, 'Column2', 'U')
        df.insert(2, 'Column3', '30010061')
        df.insert(3, 'Column4', '向日葵')
        df.drop_duplicates(keep='first', inplace=True)

        st.write("✅ Preview:")
        st.dataframe(df)

        filename = "processed_sunflower.xlsx"
        df.to_excel(filename, index=False, header=False)
        with open(filename, "rb") as f:
            st.download_button("📥 Download File", f, file_name=filename)

# ────────────────────────────── 30010013 酒田 ──────────────────────────────
elif transformation_choice == "30010013 酒田":
    raw_file = st.file_uploader("Upload Raw Sales Data", type=["xls", "xlsx"], key="sakata_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="sakata_mapping")

    if raw_file and mapping_file:
        df_raw = pd.read_excel(raw_file, sheet_name=0, header=None)

        # Date extraction
        date_string = str(df_raw.iloc[4, 0])
        match = re.search(r'至\s*(\d{3}/\d{2}/\d{2})', date_string)
        if match:
            y, m, d = map(int, match.group(1).split('/'))
            final_date = f"{y + 1911}{m:02d}{d:02d}"
        else:
            final_date = None

        data, current_product_code, current_product_name = [], None, None

        for _, row in df_raw.iterrows():
            col_a = str(row[0]).strip() if pd.notna(row[0]) else ""
            col_b = str(row[1]).strip() if pd.notna(row[1]) else ""
            col_f = row[5] if pd.notna(row[5]) else None

            if "貨品編號" in col_a and "貨品名稱" in col_a:
                match = re.search(r'貨品編號[:：]([A-Z0-9\-]+)\s+貨品名稱[:：](.+)', col_a)
                if match:
                    current_product_code = match.group(1).strip()
                    current_product_name = match.group(2).strip()
                continue

            if "小計" in col_a or "小計" in col_b:
                continue

            if col_a and col_a[0] in "DCEPM":
                if col_f and isinstance(col_f, (int, float)):
                    data.append([col_a, col_b, final_date, current_product_code, current_product_name, int(col_f)])

        df = pd.DataFrame(data, columns=["Customer Code", "Customer Name", "Date", "Product Code", "Product Name", "Quantity"])

        # Customer Mapping
        cust_map = pd.read_excel(mapping_file, sheet_name="Customer Mapping")
        cust_map = cust_map[["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"]].drop_duplicates()
        df = df.merge(cust_map, left_on="Customer Code", right_on="ASI_CRM_Offtake_Customer_No__c", how="left")
        df["Customer Code"] = df["ASI_CRM_JDE_Cust_No_Formula__c"].astype(str).str.replace(r"\.0$", "", regex=True)
        df.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"], inplace=True)

        # SKU Mapping
        sku_map = pd.read_excel(mapping_file, sheet_name="SKU Mapping")
        sku_map = sku_map[["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"]].drop_duplicates()
        df = df.merge(sku_map, left_on="Product Code", right_on="ASI_CRM_Offtake_Product__c", how="left")
        df.insert(df.columns.get_loc("Product Code"), "PRT Product Code", df["ASI_CRM_SKU_Code__c"].astype(str).str.strip())
        df.drop(columns=["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"], inplace=True)

        # Add fixed columns
        df.insert(0, "Column1", "INV")
        df.insert(1, "Column2", "U")
        df.insert(2, "Column3", "30010013")
        df.insert(3, "Column4", "酒田 ON")

        st.write("✅ Preview:")
        st.dataframe(df)

        filename = "30010013 transformation.xlsx"
        df.to_excel(filename, index=False, header=False)
        with open(filename, "rb") as f:
            st.download_button("📥 Download File", f, file_name=filename)
