import streamlit as st
import pandas as pd
import io
import re
import os

st.set_page_config(page_title="Washing Date Processor", layout="wide")
st.title("📊 Washing Date Processor")

# =========================
# SESSION STATE
# =========================
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "output" not in st.session_state:
    st.session_state.output = None

# =========================
# RESET
# =========================
if st.button("🔄 Reset"):
    st.session_state.output = None
    st.session_state.summary = None
    st.session_state.file = None
    st.session_state.uploader_key += 1
    st.rerun()

# =========================
# UPLOAD
# =========================
file1 = st.file_uploader("📂 Upload File 1 (Lot/Serial)", type=["xls", "xlsx", "csv"], key=f"file1_{st.session_state.uploader_key}")
file2 = st.file_uploader("📂 Upload File 2 (Runcard / Barcode)", type=["xls", "xlsx", "csv"], key=f"file2_{st.session_state.uploader_key}")

# =========================
# READ FUNCTIONS (คงเดิมตามโค้ดที่ใช้ได้)
# =========================
def read_excel(file):
    try:
        return pd.read_excel(file, engine="openpyxl", header=None)
    except:
        return pd.read_excel(file, engine="xlrd", header=None)

def read_file1(file):
    df = read_excel(file)
    col = 5
    start_row = 16
    data = df.iloc[start_row:, col]
    lot_list = []
    for val in data:
        if pd.isna(val): break
        val_str = str(val).strip()
        if val_str == "": break
        lot_list.append(val_str)
    return pd.DataFrame({"Lot": lot_list})

def read_file2(file):
    df = read_excel(file)
    header_row = None
    for i in range(20):
        row = df.iloc[i].astype(str).str.lower()
        if row.str.contains("runcard").any() and row.str.contains("barcode").any():
            header_row = i
            break
    if header_row is None:
        st.error("❌ หา header ไม่เจอ (Runcard / Barcode)")
        return pd.DataFrame()

    df.columns = df.iloc[header_row]
    df = df[header_row + 1:]
    df.columns = df.columns.astype(str).str.strip().str.lower()
    
    lot_cols = [c for c in df.columns if "runcard" in str(c).lower()]
    barcode_cols = [c for c in df.columns if "barcode" in str(c).lower()]

    if not lot_cols or not barcode_cols:
        st.error(f"❌ หา column ไม่เจอ")
        return pd.DataFrame()

    df_out = df[[lot_cols[0], barcode_cols[0]]].copy()
    df_out.columns = ["Lot", "Barcode No"]
    df_out = df_out.dropna(subset=["Lot"])
    df_out["Lot"] = df_out["Lot"].astype(str).str.strip()
    return df_out

# =========================
# EXTRACT LOGIC (ปรับตามโจทย์ใหม่)
# =========================
def extract_info(barcode):
    try:
        s = str(barcode).strip()
        # หาตำแหน่งตัวอักษรเพื่อเริ่มตัด WW/Day แบบเดิม
        match = re.search('[A-Za-z]', s)
        if not match: return None, None, None
        
        start = match.start()
        ww_day_code = s[start+3:start+6] # ได้ 3 หลัก เช่น 433
        
        # 🆕 ปี: นับจากท้ายสุด ตัวที่ 4
        # เช่น ...6101 -> ตัวที่ 4 นับจากท้ายคือ '6'
        year_char = s[-4] 
        year_val = 2020 + int(year_char) # 5 -> 2025, 6 -> 2026
        
        return int(ww_day_code[:2]), int(ww_day_code[2]), year_val
    except:
        return None, None, None

# =========================
# PROCESS
# =========================
if st.button("🚀 Process"):
    if file1 is None or file2 is None:
        st.warning("⚠️ กรุณาอัพโหลดไฟล์ให้ครบ")
    else:
        # 1. รวม Database 2025 + 2026
        db_frames = []
        for year_file in ["2025.txt", "2026.txt"]:
            if os.path.exists(year_file):
                # ใช้เครื่องหมายคอมม่าแยกตามมาตรฐาน csv
                tmp_db = pd.read_csv(year_file)
                # ล้างหัวคอลัมน์กันช่องว่างแฝง
                tmp_db.columns = tmp_db.columns.astype(str).str.strip()
                db_frames.append(tmp_db)
        
        if not db_frames:
            st.error("❌ ไม่พบไฟล์ฐานข้อมูล 2025.txt หรือ 2026.txt")
        else:
            date_db = pd.concat(db_frames).drop_duplicates()
            
            # อ่านไฟล์งาน
            df1 = read_file1(file1)
            df2 = read_file2(file2)

            # Merge ไฟล์ 1 และ 2
            merged = pd.merge(df1, df2, on="Lot", how="left").drop_duplicates(subset=["Lot"])

            # Extract WW, Day, Year จาก Barcode
            extracted = merged['Barcode No'].apply(lambda x: pd.Series(extract_info(x)))
            merged[['WW', 'Day', 'Year']] = extracted

            # ✅ แปลง dtype ให้ตรงกัน 100% (สำคัญมากสำหรับการ Merge)
            merged["WW"] = pd.to_numeric(merged["WW"], errors="coerce")
            merged["Day"] = pd.to_numeric(merged["Day"], errors="coerce")
            merged["Year"] = pd.to_numeric(merged["Year"], errors="coerce")

            date_db["WW"] = pd.to_numeric(date_db["WW"], errors="coerce")
            date_db["Day"] = pd.to_numeric(date_db["Day"], errors="coerce")
            date_db["Year"] = pd.to_numeric(date_db["Year"], errors="coerce")

            # ✅ MERGE 3 ขา: Year, WW, Day
            result = pd.merge(merged, date_db, on=["Year", "WW", "Day"], how="left")

            # OUTPUT
            output = result[["Lot", "Barcode No", "Year", "WW", "Day", "Date"]].copy()
            output = output.rename(columns={"Date": "Washing Date"})
            output = output[output["Lot"].astype(str).str.lower() != "lot/serial"]
            output = output.reset_index(drop=True)

            # SUMMARY
            summary = output.groupby("Washing Date")["Lot"].count().reset_index().rename(columns={"Lot": "Total Lot"})

            # SAVE TO SESSION
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                output.to_excel(writer, index=False, sheet_name="Result")
                summary.to_excel(writer, index=False, sheet_name="Summary")

            st.session_state.output = output
            st.session_state.summary = summary
            st.session_state.file = buffer.getvalue()

# =========================
# SHOW RESULT
# =========================
if "output" in st.session_state and st.session_state.output is not None:
    st.success("✅ Process สำเร็จ")
    st.subheader("📋 Result")
    st.dataframe(st.session_state.output)

    st.subheader("📊 Summary")
    st.dataframe(st.session_state.summary)

    st.download_button(
        label="📥 Download Excel",
        data=st.session_state.file,
        file_name="washing_date_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
