import streamlit as st
import pandas as pd
import io
import re
import os

st.set_page_config(page_title="Washing Date Processor", layout="wide")
st.title("📊 Washing Date Processor (Final Fix)")

# =========================
# SESSION STATE & RESET
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
    try:
        s = str(barcode).strip()
        match = re.search('[A-Za-z]', s)
        if not match: return None, None, None
        start = match.start()
        
        ww = int(s[start+3:start+5])
        day = int(s[start+5])
        # ปี: นับจากท้ายสุด ตัวที่ 4 (6 -> 2026, 5 -> 2025)
        year_digit = int(s[-4])
        year_val = 2020 + year_digit
        
        return ww, day, year_val
    except:
        return None, None, None

# =========================
# UPLOAD & PROCESS
# =========================
file1 = st.file_uploader("📂 File 1 (Lot)", type=["xls", "xlsx"], key=f"f1_{st.session_state.uploader_key}")
file2 = st.file_uploader("📂 File 2 (Barcode)", type=["xls", "xlsx"], key=f"f2_{st.session_state.uploader_key}")

if st.button("🚀 Process"):
    if file1 and file2:
        # 1. โหลดและจัดการ Database แบบยืดหยุ่น
        db_frames = []
        for y_val in [2025, 2026]:
            fname = f"{y_val}.txt"
            if os.path.exists(fname):
                tmp_db = pd.read_csv(fname)
                tmp_db.columns = tmp_db.columns.astype(str).str.strip()
                
                # ถ้าในไฟล์ไม่มีคอลัมน์ Year ให้เติมให้เองเลยตามชื่อไฟล์
                if "Year" not in tmp_db.columns:
                    tmp_db["Year"] = y_val
                
                db_frames.append(tmp_db)
        
        if not db_frames:
            st.error("❌ ไม่พบไฟล์ 2025.txt หรือ 2026.txt")
        else:
            date_db = pd.concat(db_frames).drop_duplicates()
            if "Date" in date_db.columns:
                date_db = date_db.rename(columns={"Date": "Washing Date"})

            # 2. อ่านไฟล์งาน
            df1 = read_file1(file1)
            df2 = read_file2(file2)
            merged = pd.merge(df1, df2, on="Lot", how="left").drop_duplicates(subset=["Lot"])

            # 3. ดึงข้อมูลจาก Barcode
            extracted = merged['Barcode No'].apply(lambda x: pd.Series(extract_info(x)))
            merged[['WW', 'Day', 'Year']] = extracted

            # 4. แปลงข้อมูลเป็นตัวเลขให้หมดก่อน Merge
            for col in ["WW", "Day", "Year"]:
                merged[col] = pd.to_numeric(merged[col], errors="coerce")
                date_db[col] = pd.to_numeric(date_db[col], errors="coerce")

            # 5. MERGE (ใช้ 3 ขา)
            result = pd.merge(merged, date_db, on=["Year", "WW", "Day"], how="left")

            # จัดการผลลัพธ์
            output = result[["Lot", "Barcode No", "Year", "WW", "Day", "Washing Date"]].copy()
            summary = output.dropna(subset=["Washing Date"]).groupby("Washing Date")["Lot"].count().reset_index().rename(columns={"Lot": "Total Lot"})

            st.session_state.output = output
            st.session_state.summary = summary
            
            # สร้างไฟล์ Excel
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                output.to_excel(writer, index=False, sheet_name="Result")
                summary.to_excel(writer, index=False, sheet_name="Summary")
            st.session_state.file = buf.getvalue()

# =========================
# DISPLAY
# =========================
if "output" in st.session_state and st.session_state.output is not None:
    st.success("✅ ประมวลผลสำเร็จ")
    st.dataframe(st.session_state.output, use_container_width=True)
    
    if not st.session_state.summary.empty:
        st.subheader("📊 Summary")
        st.dataframe(st.session_state.summary)
        st.download_button("📥 Download Excel", st.session_state.file, "result.xlsx")
    else:
        st.warning("⚠️ Washing Date ไม่ขึ้น: เป็นไปได้ว่าข้อมูลในไฟล์ .txt ไม่ตรงกับ WW/Day/Year ที่ดึงได้")
