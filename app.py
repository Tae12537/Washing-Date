import streamlit as st
import pandas as pd
import io
import re
import os

st.set_page_config(page_title="Washing Date Processor", layout="wide")
st.title("📊 Washing Date Processor")

# =========================
# FUNCTIONS
# =========================
def read_excel(file):
    try:
        return pd.read_excel(file, engine="openpyxl", header=None)
    except:
        return pd.read_excel(file, engine="xlrd", header=None)

def read_file1(file):
    df = read_excel(file)
    col = 5  # Column F
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
    for i in range(min(20, len(df))):
        row = df.iloc[i].astype(str).str.lower()
        if row.str.contains("runcard").any() and row.str.contains("barcode").any():
            header_row = i
            break
    if header_row is None:
        return pd.DataFrame()

    df.columns = df.iloc[header_row]
    df_data = df[header_row + 1:].copy()
    df_data.columns = df_data.columns.astype(str).str.strip().str.lower()

    lot_cols = [c for c in df_data.columns if "runcard" in str(c)]
    barcode_cols = [c for c in df_data.columns if "barcode" in str(c)]
    packing_cols = [c for c in df_data.columns if "packing" in str(c) or str(c) == "q4"]

    if not lot_cols or not barcode_cols:
        return pd.DataFrame()

    df_out = df_data[[lot_cols[0], barcode_cols[0]] + ([packing_cols[0]] if packing_cols else [])].copy()
    new_names = {lot_cols[0]: "Lot", barcode_cols[0]: "Barcode No"}
    if packing_cols: new_names[packing_cols[0]] = "Packing Date"
    
    df_out = df_out.rename(columns=new_names)
    df_out = df_out.dropna(subset=["Lot"])
    df_out["Lot"] = df_out["Lot"].astype(str).str.strip()
    if "Packing Date" in df_out.columns:
        df_out["Packing Date"] = pd.to_datetime(df_out["Packing Date"], dayfirst=True, errors='coerce')
    return df_out

def extract_ww_day(barcode):
    try:
        s = str(barcode)
        match = re.search('[A-Za-z]', s)
        if not match: return None, None
        start = match.start()
        code = s[start+3:start+6]
        if len(code) != 3 or not code.isdigit(): return None, None
        return int(code[:2]), int(code[2])
    except: return None, None

# =========================
# UI
# =========================
if "uploader_key" not in st.session_state: st.session_state.uploader_key = 0

file1 = st.file_uploader("📂 File 1", type=["xls", "xlsx", "csv"], key=f"f1_{st.session_state.uploader_key}")
file2 = st.file_uploader("📂 File 2", type=["xls", "xlsx", "csv"], key=f"f2_{st.session_state.uploader_key}")

col1, col2, _ = st.columns([1, 1, 4])
with col1: btn_process = st.button("🚀 Process", use_container_width=True)
with col2:
    if st.button("🔄 Reset", use_container_width=True):
        st.session_state.uploader_key += 1
        st.rerun()

# =========================
# LOGIC
# =========================
if btn_process and file1 and file2:
    df1 = read_file1(file1)
    df2 = read_file2(file2)

    if not df1.empty and not df2.empty:
        # กำหนดปี
        detected_year = 2026
        if "Packing Date" in df2.columns and not df2["Packing Date"].isnull().all():
            val = df2["Packing Date"].dropna()
            if not val.empty:
                detected_year = pd.to_datetime(val.iloc[0]).year
        
        db_file = f"{int(detected_year)}.txt"
        
        if os.path.exists(db_file):
            # 1. โหลด DB แบบปลอดภัย
            date_db = pd.read_csv(db_file, skipinitialspace=True)
            date_db.columns = [c.strip() for c in date_db.columns]
            
            # บังคับชื่อคอลัมน์ Date เป็น Washing Date
            date_db = date_db.rename(columns={"Date": "Washing Date"})
            
            # แปลง Year, WW, Day เป็นตัวเลขแบบปลอดภัย (แถวไหนไม่ใช่เลขจะกลายเป็น NaN แล้วถูกลบ)
            for col in ["Year", "WW", "Day"]:
                if col in date_db.columns:
                    date_db[col] = pd.to_numeric(date_db[col], errors='coerce')
            
            date_db = date_db.dropna(subset=["Year", "WW", "Day"])
            date_db[["Year", "WW", "Day"]] = date_db[["Year", "WW", "Day"]].astype(int)

            # 2. เตรียมไฟล์งาน
            merged = pd.merge(df1, df2, on="Lot", how="left").drop_duplicates(subset=["Lot"])
            merged[['WW', 'Day']] = merged['Barcode No'].apply(lambda x: pd.Series(extract_ww_day(x)))
            
            merged["WW"] = pd.to_numeric(merged["WW"], errors="coerce").fillna(0).astype(int)
            merged["Day"] = pd.to_numeric(merged["Day"], errors="coerce").fillna(0).astype(int)
            merged["Year"] = int(detected_year)

            # 3. Merge
            final = pd.merge(merged, date_db, on=["Year", "WW", "Day"], how="left")

            # 4. ผลลัพธ์
            cols_show = ["Lot", "Barcode No", "Year", "WW", "Day", "Washing Date", "Packing Date"]
            output = final[[c for c in cols_show if c in final.columns]].copy()
            if "Packing Date" in output.columns:
                output["Packing Date"] = pd.to_datetime(output["Packing Date"]).dt.date

            # 5. สรุป
            summary = pd.DataFrame()
            if "Washing Date" in output.columns:
                valid_data = output.dropna(subset=["Washing Date"])
                if not valid_data.empty:
                    summary = valid_data.groupby("Washing Date").size().reset_index(name="Total Lot")

            # แสดงผล
            st.success(f"✅ ประมวลผลสำเร็จ (ฐานข้อมูล: {db_file})")
            st.dataframe(output, use_container_width=True)
            
            if not summary.empty:
                st.subheader("📊 Summary")
                st.dataframe(summary, use_container_width=True)
                
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as wr:
                    output.to_excel(wr, index=False, sheet_name="Result")
                    summary.to_excel(wr, index=False, sheet_name="Summary")
                st.download_button("📥 Download Excel", buf.getvalue(), "report.xlsx")
            else:
                st.warning("⚠️ ไม่พบ Washing Date ที่ตรงกับฐานข้อมูล (ตรวจสอบ Year, WW, Day ในไฟล์ .txt)")
        else:
            st.error(f"❌ ไม่พบไฟล์ฐานข้อมูล `{db_file}` ในระบบ")
