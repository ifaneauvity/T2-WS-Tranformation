import streamlit as st
import pandas as pd
import re

st.title("📊 WS Transformation")
st.write("Upload an Excel file and choose the transformation format.")

transformation_choice = st.radio("Select Transformation Format:", [
    "宏酒樽", "向日葵", "30010010 酒倉盛豐行", "30010013 酒田"
])

if transformation_choice == "30010010 酒倉盛豐行":
    raw_data_file = st.file_uploader("Upload Raw Sales Data", type=["xlsx"], key="ws_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="ws_mapping")

    if raw_data_file and mapping_file:
        raw_df = pd.read_excel(raw_data_file, sheet_name="給廠商", header=None)

        date_string = str(raw_df.iloc[4, 0])
        match = re.search(r'至\s*(\d{3}/\d{2}/\d{2})', date_string)
        final_date = f"{int(match.group(1).split('/')[0]) + 1911}{int(match.group(1).split('/')[1]):02d}{int(match.group(1).split('/')[2]):02d}" if match else None

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
                data.append([col_a, col_b, final_date, current_product_code, current_product_name, int(col_d)])

        df_cleaned = pd.DataFrame(data, columns=["Customer Code", "Customer Name", "Date", "Product Code", "Product Name", "Quantity"])

        mapping_customer = pd.read_excel(mapping_file, sheet_name="Customer Mapping")
        mapping_customer = mapping_customer[[
            "ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Customer_No__c")

        df_cleaned = df_cleaned.merge(mapping_customer, left_on="Customer Code", right_on="ASI_CRM_Offtake_Customer_No__c", how="left")
        df_cleaned["Customer Code"] = df_cleaned["ASI_CRM_JDE_Cust_No_Formula__c"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        df_cleaned.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"], inplace=True)

        mapping_sku = pd.read_excel(mapping_file, sheet_name="SKU Mapping")
        mapping_sku = mapping_sku[["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"]].drop_duplicates(subset="ASI_CRM_Offtake_Product__c")

        df_cleaned = df_cleaned.merge(mapping_sku, left_on="Product Code", right_on="ASI_CRM_Offtake_Product__c", how="left")
        df_cleaned.insert(df_cleaned.columns.get_loc("Product Code"), "PRT Product Code", df_cleaned["ASI_CRM_SKU_Code__c"].astype(str).str.strip())
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
    raw_data_file = st.file_uploader("Upload Raw Sales Data", type=["xls"], key="sakata_raw")
    mapping_file = st.file_uploader("Upload Mapping File", type=["xlsx"], key="sakata_mapping")

    if raw_data_file and mapping_file:
        raw_df = pd.read_excel(raw_data_file, sheet_name=0, header=None)

        date_string = str(raw_df.iloc[4, 0])
        match = re.search(r'至\s*(\d{3}/\d{2}/\d{2})', date_string)
        final_date = f"{int(match.group(1).split('/')[0]) + 1911}{int(match.group(1).split('/')[1]):02d}{int(match.group(1).split('/')[2]):02d}" if match else None

        current_product_code = None
        current_product_name = None
        data = []

        for _, row in raw_df.iterrows():
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

            if col_a and col_a[0] in "DCEMP" and col_f and isinstance(col_f, (int, float)):
                data.append([col_a, col_b, final_date, current_product_code, current_product_name, int(col_f)])

        df_cleaned = pd.DataFrame(data, columns=["Customer Code", "Customer Name", "Date", "Product Code", "Product Name", "Quantity"])

        mapping_customer = pd.read_excel(mapping_file, sheet_name="Customer Mapping")
        mapping_customer = mapping_customer[[
            "ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"
        ]].drop_duplicates(subset="ASI_CRM_Offtake_Customer_No__c")

        df_cleaned = df_cleaned.merge(mapping_customer, left_on="Customer Code", right_on="ASI_CRM_Offtake_Customer_No__c", how="left")
        df_cleaned["Customer Code"] = df_cleaned["ASI_CRM_JDE_Cust_No_Formula__c"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        df_cleaned.drop(columns=["ASI_CRM_Offtake_Customer_No__c", "ASI_CRM_JDE_Cust_No_Formula__c"], inplace=True)

        mapping_sku = pd.read_excel(mapping_file, sheet_name="SKU Mapping")
        mapping_sku = mapping_sku[["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"]].drop_duplicates(subset="ASI_CRM_Offtake_Product__c")

        df_cleaned = df_cleaned.merge(mapping_sku, left_on="Product Code", right_on="ASI_CRM_Offtake_Product__c", how="left")
        df_cleaned.insert(df_cleaned.columns.get_loc("Product Code"), "PRT Product Code", df_cleaned["ASI_CRM_SKU_Code__c"].astype(str).str.strip())
        df_cleaned.drop(columns=["ASI_CRM_Offtake_Product__c", "ASI_CRM_SKU_Code__c"], inplace=True)

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
