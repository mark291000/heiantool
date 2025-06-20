import streamlit as st
import pdfplumber
import pandas as pd
import io
import os
import re
from tempfile import NamedTemporaryFile

st.set_page_config(page_title="PDF Extractor Tool", layout="wide")
st.title("HEIAN Table Extractor Tool")
st.markdown("📌 For any issues related to the app, please contact Mark Dang.")

standard_columns = [
    "Part ID",
    "Part Name",
    "Cart Loading",
    "Qty Req",
    "Qty Nested",
    "Part Description",
    "Production Instructions",
    "Material"
]

def clean_and_align_table(df_raw):
    df_raw = df_raw.dropna(how="all").reset_index(drop=True)

    def is_col_empty_or_zero(col):
        all_none = col.isna().all()
        try:
            all_zero = (col.fillna(0).astype(float) == 0).all()
        except:
            all_zero = False
        return all_none or all_zero

    df_raw = df_raw[[col for col in df_raw.columns if not is_col_empty_or_zero(df_raw[col])]]
    n_col = df_raw.shape[1]

    if n_col == 8:
        df_raw.columns = standard_columns
    elif n_col == 7:
        temp_cols = [col for col in standard_columns if col != "Cart Loading"]
        df_raw.columns = temp_cols
        df_raw.insert(2, "Cart Loading", pd.NA)
    else:
        raise ValueError(f"❌ Bảng có {n_col} cột. Yêu cầu 7 hoặc 8 cột.")

    return df_raw

def extract_data_from_pdf(file_bytes, filename):
    all_tables = []
    base_name = os.path.splitext(filename)[0]

    with NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(file_bytes.getvalue())
        tmp_path = tmp.name

    with pdfplumber.open(tmp_path) as pdf:
        full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])
        match = re.search(r"(\d+(\.\d+)?)\s*Sheet\(s\)\s*=\s*(\d+(\.\d+)?)\s*Kit\(s\)", full_text, re.IGNORECASE)
        sheet_count = float(match.group(1)) if match else None
        kit_count = float(match.group(3)) if match else None

        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables:
                continue
            for table in tables:
                if not table or len(table) < 2:
                    continue

                data_rows = table[1:]
                df_temp = pd.DataFrame(data_rows)
                df_temp = df_temp[~df_temp.apply(lambda row: row.astype(str).str.contains("Yield:", case=False).any(), axis=1)]

                if df_temp.empty:
                    continue

                try:
                    df_clean = clean_and_align_table(df_temp)
                    df_clean.insert(1, "Program", base_name)
                    df_clean["Sheet"] = sheet_count
                    df_clean["Kit"] = kit_count
                    all_tables.append(df_clean)
                except Exception as e:
                    st.warning(f"⚠️ Lỗi khi xử lý bảng từ {filename}: {e}")

    return pd.concat(all_tables, ignore_index=True) if all_tables else pd.DataFrame()

uploaded_files = st.file_uploader("📂 Kéo và thả file PDF vào đây", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    df_list = []
    total = len(uploaded_files)
    progress = st.progress(0)
    status = st.empty()

    for idx, file in enumerate(uploaded_files, 1):
        status.text(f"🔍 Đang xử lý: {file.name} ({idx}/{total})")
        df = extract_data_from_pdf(file, file.name)
        if not df.empty:
            df_list.append(df)
        progress.progress(idx / total)

    if df_list:
        combined_df = pd.concat(df_list, ignore_index=True)
        combined_df = combined_df[combined_df["Part Name"].notna()]

        for col in ["Qty Req", "Qty Nested", "Sheet", "Kit"]:
            combined_df[col] = pd.to_numeric(combined_df[col], errors="coerce").fillna(0)

        grouped_df = combined_df.groupby(["Part Name", "Program"], dropna=False).agg({
            "Sheet": "first",
            "Kit": "first",
            "Part ID": "first",
            "Cart Loading": "first",
            "Qty Req": "sum",
            "Qty Nested": "sum",
            "Part Description": "first",
            "Production Instructions": "first",
            "Material": "first",
        }).reset_index()

        cols = grouped_df.columns.tolist()
        cols.insert(0, cols.pop(cols.index("Program")))
        grouped_df = grouped_df[cols]
        grouped_df["Part ID"] = pd.to_numeric(grouped_df["Part ID"], errors="coerce").fillna(0).astype(int)

        grouped_df["Usage Wood Gross"] = grouped_df.apply(
            lambda row: round(32.96 * row["Sheet"] / row["Kit"], 3) if row["Kit"] else None, axis=1
        )
        grouped_df["Usage Wood Net"] = grouped_df.apply(
            lambda row: round(32 * row["Sheet"] / row["Kit"], 3) if row["Kit"] else None, axis=1
        )
        grouped_df["Usage CNC Part"] = grouped_df.apply(
            lambda row: row["Qty Nested"] if "offal" in str(row["Part Name"]).lower()
            else round(row["Qty Nested"] / row["Kit"], 3) if row["Kit"] else None, axis=1
        )

        grouped_df = grouped_df.sort_values(by=["Program", "Part Name"], ignore_index=True)

        st.success("✅ Hoàn tất xử lý!")
        st.dataframe(grouped_df, use_container_width=True)

        # Export file Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            grouped_df.to_excel(writer, index=False, sheet_name="Summary")
        st.download_button(
            label="📥 Tải Excel kết quả",
            data=output.getvalue(),
            file_name="extracted_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("❌ Không tìm thấy dữ liệu hợp lệ.")
