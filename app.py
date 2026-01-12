import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from tempfile import NamedTemporaryFile

st.set_page_config(page_title="HEIAN OFFAL Extractor", page_icon="📊", layout="wide")

st.title("📊 HEIAN Table Extractor - OFFAL Filter")
st.markdown("---")

# Cột chuẩn
standard_columns = [
    "Part ID", "Part Name", "Cart Loading", "Qty Req", 
    "Qty Nested", "Part Description", "Production Instructions", "Material"
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

def extract_thickness_and_scrap_from_text(text):
    thickness = None
    scrap = None
    
    thickness_match = re.search(r'Thickness:\s*(\d+(?:\.\d+)?)\s*MM', text, re.IGNORECASE)
    if thickness_match:
        thickness = float(thickness_match.group(1))
    
    scrap_match = re.search(r'Scrap:\s*(\d+(?:\.\d+)?)\s*%', text, re.IGNORECASE)
    if scrap_match:
        scrap = scrap_match.group(1) + "%"
    
    return thickness, scrap

def get_material_code(thickness):
    """Xác định mã Material dựa trên Thickness"""
    if thickness == 15:
        return "280040WNK"
    elif thickness == 18:
        return "280045WNK"
    elif thickness == 12:
        return "280090WNK"
    elif thickness == 9:
        return "280062"
    else:
        return ""

def extract_data_from_pdf(file_bytes, filename):
    all_tables = []
    base_name = filename.replace('.pdf', '')
    
    thickness_value = None
    scrap_sheet1 = None
    scrap_sheet2 = None

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])
        match = re.search(r"(\d+(\.\d+)?)\s*Sheet\(s\)\s*=\s*(\d+(\.\d+)?)\s*Kit\(s\)", full_text, re.IGNORECASE)
        sheet_count = float(match.group(1)) if match else None
        kit_count = float(match.group(3)) if match else None

        for page_num, page in enumerate(pdf.pages, 1):
            page_text = page.extract_text() or ""
            thickness, scrap = extract_thickness_and_scrap_from_text(page_text)
            
            if thickness and thickness_value is None:
                thickness_value = thickness
            
            if scrap:
                if page_num == 1:
                    scrap_sheet1 = scrap
                elif page_num == 2:
                    scrap_sheet2 = scrap
            
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
                    df_clean.insert(0, "Program", base_name)
                    df_clean["Sheet"] = sheet_count
                    df_clean["Kit"] = kit_count
                    df_clean["Thickness"] = thickness_value
                    df_clean["Scrap Sheet1"] = scrap_sheet1
                    df_clean["Scrap Sheet2"] = scrap_sheet2
                    all_tables.append(df_clean)
                except Exception as e:
                    st.warning(f"⚠️ Lỗi khi xử lý bảng từ {filename}: {e}")

    return pd.concat(all_tables, ignore_index=True) if all_tables else pd.DataFrame()

# Upload files
uploaded_files = st.file_uploader(
    "📂 Chọn hoặc kéo thả file PDF", 
    type=['pdf'], 
    accept_multiple_files=True
)

if uploaded_files:
    with st.spinner('🔄 Đang xử lý các file PDF...'):
        df_list = []
        
        progress_bar = st.progress(0)
        total = len(uploaded_files)
        
        for idx, uploaded_file in enumerate(uploaded_files):
            st.info(f"🔍 Đang xử lý: {uploaded_file.name} ({idx + 1}/{total})")
            
            file_bytes = uploaded_file.read()
            df = extract_data_from_pdf(file_bytes, uploaded_file.name)
            
            if not df.empty:
                df_list.append(df)
            
            progress_bar.progress((idx + 1) / total)

    if df_list:
        combined_df = pd.concat(df_list, ignore_index=True)
        combined_df = combined_df[combined_df["Part Name"].notna()]

        # Lọc chỉ lấy OFFAL
        combined_df = combined_df[combined_df["Part Name"].str.contains("OFFAL", case=False, na=False)]

        if combined_df.empty:
            st.error("❌ Không tìm thấy Part name nào chứa 'OFFAL'")
        else:
            # Lấy OFFAL đầu tiên theo Program
            combined_df = combined_df.groupby("Program").first().reset_index()

            # Chuyển đổi kiểu dữ liệu
            for col in ["Qty Req", "Qty Nested", "Sheet", "Kit"]:
                if col in combined_df.columns:
                    combined_df[col] = pd.to_numeric(combined_df[col], errors="coerce").fillna(0)

            # Tạo cột Block Offal
            combined_df["Block Offal"] = combined_df["Qty Nested"]
            
            # Tạo cột Material dựa trên Thickness
            combined_df["Material"] = combined_df["Thickness"].apply(get_material_code)

            # Chỉ hiển thị các cột yêu cầu
            final_columns = ["Program", "Sheet", "Kit", "Thickness", "Scrap Sheet1", "Scrap Sheet2", "Block Offal", "Material"]
            result_df = combined_df[final_columns]

            # Sắp xếp theo Program
            result_df = result_df.sort_values(by=["Program"], ignore_index=True)

            st.success(f"✅ Hoàn tất! Tổng số dòng OFFAL: {len(result_df)}")
            
            # Hiển thị bảng
            st.dataframe(result_df, use_container_width=True)

            # Tạo file Excel để download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                result_df.to_excel(writer, index=False, sheet_name="OFFAL Parts")
            
            output.seek(0)
            
            st.download_button(
                label="📥 Download Excel",
                data=output,
                file_name="extracted_offal_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("❌ Không tìm thấy dữ liệu hợp lệ trong các file PDF.")
else:
    st.info("👆 Vui lòng upload file PDF để bắt đầu")
