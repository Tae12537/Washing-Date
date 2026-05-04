import streamlit as st
import pandas as pd
import io
import re
import os

st.set_page_config(page_title="Washing Date Processor", layout="wide")
st.title("📊 Washing Date Processor (Database Edition)")

# =========================
# SESSION STATE
# =========================
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

if st.button("🔄 Reset"):
    st.session_state.output = None
    st.session_state.uploader_key += 1
    st.rerun()

# =========================
# READ FUNCTIONS
# =========================
def read_excel(file):
    try:
        return pd.read_excel(file, engine="openpyxl", header=None)
    except:
        return pd.read_excel(file, engine="xlrd", header=None)

def read_file1(file):
    df = read_excel(file)
    data = df.iloc[16:, 5] # Column F, Row 17+
    lot_list = [str(val).strip() for val in data if pd.notna(val) and str(val).strip() != ""]
    return pd.DataFrame({"Lot": lot_list})

def read_file2(file):
    df = read_excel(file)
    header_row = None
    for i in range(20):
        row = df.iloc[i].astype(str).str.lower()
        if row.str.contains("runcard").any() and row.str.contains("barcode").any():
            header_row = i
            break
    if header_row is None: return pd.DataFrame()

    df.columns = df.iloc[header_row]
    df = df[header_row + 1:]
    df.columns = df.columns.astype(str).str.strip().str.lower()
    
    lot_cols = [c for c in df.columns if "runcard" in str(c).lower()]
    barcode_cols = [c for c in df.columns if "barcode" in str(c).lower()]

    if not lot_cols or not barcode_cols: return pd.DataFrame()

    df_out = df[[lot_cols[0], barcode_cols[0]]].copy()
    df_out.columns = ["Lot", "Barcode No"]
    df_out["Lot"] = df_out["Lot"].astype(str).str.strip()
    return df_out

def extract_info(barcode):
    """ดึงข้อมูลจาก Barcode: WW(ตัวที่ 4-5), Day(ตัวที่ 6), Year(ตัวที่ 4 จากท้าย)"""
    try:
        s = str(barcode).strip()
        match = re.search('[A-Za-z]', s)
        if not match: return None, None, None
        
        start = match.start()
        ww = int(s[start+3:start+5])
        day = int(s[start+5])
        
        # ปี: ตัวที่ 4 นับจากท้ายสุด (6 -> 2026, 5 -> 2025)
        year_digit = int(s[-4])
        year_val = 2020 + year_digit
        
        return ww, day, year_val
    except:
        return None, None, None

# =========================
# UPLOAD
# =========================
file1 = st.file_uploader("📂 File 1 (Lot)", type=["xls", "xlsx"], key=f"f1_{st.session_state.uploader_key}")
file2 = st.file_uploader("📂 File 2 (Barcode)", type=["xls", "xlsx"], key=f"f2_{st.session_state.uploader_key}")

if st.button("🚀 Process"):
    if file1 and file2:
        # 1. โหลด database.txt
        db_path = "database.txt"
        if not os.path.exists(db_path):
            st.error(f"❌ ไม่พบไฟล์ `{db_path}` กรุณารวมไฟล์ 2025 และ 2026 เป็นไฟล์เดียวแล้วตั้งชื่อนี้")
        else:
            # อ่าน DB และ Clean หัวตาราง
            date_db = pd.read_csv(db_path)
            date_db.columns = date_db.columns.astype(str).str.strip()
            
            # บังคับชื่อคอลัมน์ผลลัพธ์
            if "Date" in date_db.columns:
                date_db = date_db.rename(columns={"Date": "Washing Date"})

            # 2. อ่านไฟล์งาน
            df1 = read_file1(file1)
            df2 = read_file2(file2)
            merged = pd.merge(df1, df2, on="Lot", how="left").drop_duplicates(subset=["Lot"])

            # 3. ดึงข้อมูลจาก Barcode
            extracted = merged['Barcode No'].apply(lambda x: pd.Series(extract_info(x)))
            merged[['WW', 'Day', 'Year']] = extracted

            # 4. 🚨 หัวใจสำคัญ: แปลงเป็นตัวเลข (Int) ทั้ง 2 ฝั่งเพื่อป้องกัน Error ในการ Merge
            cols_to_match = ["Year", "WW", "Day"]
            for col in cols_to_match:
                # ฝั่งไฟล์งาน
                merged[col] = pd.to_numeric(merged[col], errors="coerce")
                # ฝั่ง Database
                date_db[col] = pd.to_numeric(date_db[col], errors="coerce")

            # ลบแถวที่ดึงเลขไม่ได้ทิ้งไปก่อน
            merged = merged.dropna(subset=cols_to_match)
            date_db = date_db.dropna(subset=cols_to_match)
            
            # บังคับเป็น Int
            merged[cols_to_match] = merged[cols_to_match].astype(int)
            date_db[cols_to_match] = date_db[cols_to_match].astype(int)

            # 5. MERGE
            final_result = pd.merge(merged, date_db, on=["Year", "WW", "Day"], how="left")

            # 6. แสดงผลและเตรียมดาวน์โหลด
            output = final_result[["Lot", "Barcode No", "Year", "WW", "Day", "Washing Date"]].copy()
            summary = output.dropna(subset=["Washing Date"]).groupby("Washing Date")["Lot"].count().reset_index().rename(columns={"Lot": "Total Lot"})

            st.session_state.output = output
            st.session_state.summary = summary
            
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                output.to_excel(writer, index=False, sheet_name="Result")
                if not summary.empty:
                    summary.to_excel(writer, index=False, sheet_name="Summary")
            st.session_state.file = buf.getvalue()

# =========================
# DISPLAY
# =========================
if "output" in st.session_state and st.session_state.output is not None:
    st.success("✅ ประมวลผลสำเร็จโดยใช้ database.txt")
    st.dataframe(st.session_state.output, use_container_width=True)
    
    if not st.session_state.summary.empty:
        st.subheader("📊 Summary")
        st.dataframe(st.session_state.summary)
        st.download_button("📥 Download Excel", st.session_state.file, "washing_report.xlsx")
    else:
        st.warning("⚠️ Washing Date ยังไม่ขึ้น! กรุณาตรวจสอบข้อมูลใน database.txt ว่ามีเลข Year, WW, Day ที่ตรงกับไฟล์งานหรือไม่")
