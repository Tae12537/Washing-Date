import streamlit as st
import pandas as pd
import io
import re
import os
from datetime import datetime

st.set_page_config(page_title="Smart Washing Date Processor", layout="wide")
st.title("🚀 Smart Washing Date Processor")

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

# =========================
# UPLOAD SECTION
# =========================
col_u1, col_u2 = st.columns(2)
with col_u1:
    file1 = st.file_uploader("📂 File 1 (Lot/Serial)", type=["xls", "xlsx"], key=f"f1_{st.session_state.uploader_key}")
with col_u2:
    file2 = st.file_uploader("📂 File 2 (Runcard / Barcode)", type=["xls", "xlsx"], key=f"f2_{st.session_state.uploader_key}")

# =========================
# HELPER FUNCTIONS
# =========================
def read_excel(file):
    try:
        return pd.read_excel(file, engine="openpyxl", header=None)
    except:
        return pd.read_excel(file, engine="xlrd", header=None)

def read_file1(file):
    df = read_excel(file)
    # ดึง Lot (Col F / Index 5)
    lot_data = df.iloc[16:, 5]
    # ดึง Pack Date (Col Q / Index 16) - สมมติว่าเป็น Q4 ตามที่แจ้ง (หรือ Index 16)
    # ถ้าคำว่า Q4 หมายถึง Column ที่ 17 (Q)
    pack_date_data = df.iloc[16:, 16] 
    
    lots = []
    p_dates = []
    for l, p in zip(lot_data, pack_date_data):
        if pd.isna(l) or str(l).strip() == "": break
        lots.append(str(l).strip())
        p_dates.append(p)
        
    return pd.DataFrame({"Lot": lots, "Pack Date": p_dates})

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
    
    lot_col = [c for c in df.columns if "runcard" in c][0]
    barcode_col = [c for c in df.columns if "barcode" in c][0]
    
    df_out = df[[lot_col, barcode_col]].copy()
    df_out.columns = ["Lot", "Barcode No"]
    return df_out.dropna(subset=["Lot"])

def extract_ww_day(barcode):
    try:
        s = str(barcode)
        match = re.search('[A-Za-z]', s)
        if not match: return None, None
        start = match.start()
        code = s[start+3:start+6]
        return int(code[:2]), int(code[2])
    except: return None, None

# =========================
# PROCESS & RESET BUTTONS
# =========================
c1, c2, c3 = st.columns([1, 1, 4])
with c1: process_clicked = st.button("🚀 Process", use_container_width=True)
with c2: 
    if st.button("🔄 Reset", use_container_width=True):
        st.session_state.clear()
        st.rerun()

# =========================
# MAIN LOGIC
# =========================
if process_clicked:
    if not file1 or not file2:
        st.warning("⚠️ กรุณาอัพโหลดไฟล์ให้ครบ")
    else:
        # 1. Load All Database Files (.txt)
        all_db = []
        for f in os.listdir("."):
            if f.endswith(".txt"):
                try:
                    tdf = pd.read_csv(f)
                    tdf.columns = tdf.columns.str.strip()
                    # แปลง Date ใน DB ให้เป็น datetime
                    tdf['Date_DT'] = pd.to_datetime(tdf['Date'], format='%d-%b-%Y', errors='coerce')
                    all_db.append(tdf)
                except: pass
        
        if not all_db:
            st.error("❌ ไม่พบไฟล์ Database (.txt) ในโฟลเดอร์")
            st.stop()
            
        db_full = pd.concat(all_db, ignore_index=True)

        # 2. Read Files
        df1 = read_file1(file1)
        df2 = read_file2(file2)
        
        # Merge Lot กับ Barcode
        merged = pd.merge(df1, df2, on="Lot", how="left").drop_duplicates(subset=["Lot"])
        merged[['WW', 'Day']] = merged['Barcode No'].apply(lambda x: pd.Series(extract_ww_day(x)))
        
        # 3. Smart Matching Logic (หา Date ที่ใกล้ Pack Date ที่สุด)
        final_results = []
        
        for _, row in merged.iterrows():
            ww, day, p_date = row['WW'], row['Day'], row['Pack Date']
            
            # กรอง DB เฉพาะ WW และ Day ที่ตรงกัน (อาจจะได้หลายบรรทัดจากหลายปี)
            matches = db_full[(db_full['WW'] == ww) & (db_full['Day'] == day)].copy()
            
            best_date = None
            if not matches.empty:
                if pd.notna(p_date):
                    # แปลง Pack Date ให้เป็น datetime เพื่อคำนวณระยะห่าง
                    p_date_dt = pd.to_datetime(p_date, errors='coerce')
                    if pd.notna(p_date_dt):
                        # หาบรรทัดที่ Date ห่างจาก Pack Date น้อยที่สุด
                        matches['diff'] = (matches['Date_DT'] - p_date_dt).abs()
                        best_date = matches.sort_values('diff').iloc[0]['Date']
                    else:
                        best_date = matches.iloc[0]['Date']
                else:
                    best_date = matches.iloc[0]['Date']
            
            final_results.append({
                "Lot": row['Lot'],
                "Barcode No": row['Barcode No'],
                "Pack Date": p_date,
                "WW": ww,
                "Day": day,
                "Washing Date": best_date
            })

        output_df = pd.DataFrame(final_results)
        
        # 4. Summary & Display
        st.success("✅ ประมวลผลสำเร็จ (ใช้ระบบ Smart Match กับ Pack Date)")
        st.dataframe(output_df, use_container_width=True)
        
        summary = output_df.groupby("Washing Date")["Lot"].count().reset_index(name="Total Lot")
        st.subheader("📊 Summary")
        st.table(summary)

        # Download
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            output_df.to_excel(writer, index=False, sheet_name="Result")
            summary.to_excel(writer, index=False, sheet_name="Summary")
        
        st.download_button("📥 Download Result", buffer.getvalue(), "result.xlsx", "application/vnd.ms-excel")
