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
# FUNCTIONS
# =========================
def read_excel(file):
    try:
        return pd.read_excel(file, engine="openpyxl", header=None)
    except:
        return pd.read_excel(file, engine="xlrd", header=None)

def read_file1(file):
    df = read_excel(file)
    col = 5  # คอลัมน์ F
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
        st.error("❌ หา header ไม่เจอ (ต้องมี Runcard และ Barcode)")
        return pd.DataFrame()

    df.columns = df.iloc[header_row]
    df_data = df[header_row + 1:].copy()
    df_data.columns = df_data.columns.astype(str).str.strip().str.lower()

    lot_cols = [c for c in df_data.columns if "runcard" in str(c)]
    barcode_cols = [c for c in df_data.columns if "barcode" in str(c)]
    packing_cols = [c for c in df_data.columns if "packing" in str(c) or str(c) == "q4"]

    if not lot_cols or not barcode_cols:
        return pd.DataFrame()

    use_cols = [lot_cols[0], barcode_cols[0]]
    if packing_cols:
        use_cols.append(packing_cols[0])

    df_out = df_data[use_cols].copy()
    new_names = {lot_cols[0]: "Lot", barcode_cols[0]: "Barcode No"}
    if packing_cols:
        new_names[packing_cols[0]] = "Packing Date"
    
    df_out = df_out.rename(columns=new_names)
    df_out = df_out.dropna(subset=["Lot"])
    df_out["Lot"] = df_out["Lot"].astype(str).str.strip()
    
    if "Packing Date" in df_out.columns:
        # ตัดเวลาออกเหลือแต่วันที่
        df_out["Packing Date"] = pd.to_datetime(df_out["Packing Date"], dayfirst=True, errors='coerce').dt.date
        
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
file1 = st.file_uploader("📂 Upload File 1 (Lot/Serial)", type=["xls", "xlsx", "csv"], key=f"f1_{st.session_state.uploader_key}")
file2 = st.file_uploader("📂 Upload File 2 (Runcard / Barcode)", type=["xls", "xlsx", "csv"], key=f"f2_{st.session_state.uploader_key}")

col1, col2, _ = st.columns([1, 1, 4])
with col1:
    btn_process = st.button("🚀 Process", use_container_width=True)
with col2:
    if st.button("🔄 Reset", use_container_width=True):
        st.session_state.output = None
        st.session_state.summary = None
        st.session_state.file = None
        st.session_state.uploader_key += 1
        st.rerun()

# =========================
# LOGIC
# =========================
if btn_process:
    if not file1 or not file2:
        st.warning("⚠️ กรุณาอัพโหลดไฟล์ให้ครบ")
    else:
        df1 = read_file1(file1)
        df2 = read_file2(file2)

        if not df1.empty and not df2.empty:
            # ตรวจสอบปีเพื่อเลือกไฟล์
            detected_year = 2026
            if "Packing Date" in df2.columns and not df2["Packing Date"].isnull().all():
                first_date = pd.to_datetime(df2["Packing Date"].dropna().iloc[0])
                detected_year = first_date.year
            
            db_filename = f"{int(detected_year)}.txt"
            
            if not os.path.exists(db_filename):
                st.error(f"❌ ไม่พบไฟล์ฐานข้อมูล `{db_filename}`")
            else:
                # 1. โหลด DB และทำความสะอาดคอลัมน์
                date_db = pd.read_csv(db_filename, skipinitialspace=True)
                date_db.columns = [c.strip() for c in date_db.columns]
                
                # มั่นใจว่าคอลัมน์ Date ชื่อตรงกันแน่นอน
                if "Date" in date_db.columns:
                    date_db = date_db.rename(columns={"Date": "Washing Date"})
                
                # บังคับประเภทข้อมูล Key ให้เป็น Int
                for col in ["Year", "WW", "Day"]:
                    if col in date_db.columns:
                        date_db[col] = pd.to_numeric(date_db[col], errors='coerce').fillna(0).astype(int)

                # 2. เตรียมไฟล์หลัก
                merged = pd.merge(df1, df2, on="Lot", how="left").drop_duplicates(subset=["Lot"])
                merged[['WW', 'Day']] = merged['Barcode No'].apply(lambda x: pd.Series(extract_ww_day(x)))
                
                # บังคับประเภทข้อมูล Key ฝั่งนี้ให้เป็น Int เช่นกัน
                merged["WW"] = pd.to_numeric(merged["WW"], errors="coerce").fillna(0).astype(int)
                merged["Day"] = pd.to_numeric(merged["Day"], errors="coerce").fillna(0).astype(int)
                
                if "Packing Date" in merged.columns:
                    merged["Year"] = pd.to_datetime(merged["Packing Date"]).dt.year.fillna(detected_year).astype(int)
                else:
                    merged["Year"] = int(detected_year)

                # 3. Merge ข้อมูล (ใช้ Year, WW, Day เป็นตัวเชื่อม)
                result = pd.merge(merged, date_db, on=["Year", "WW", "Day"], how="left")

                # กรองเฉพาะคอลัมน์ที่ต้องการแสดง
                final_cols = ["Lot", "Barcode No", "Year", "WW", "Day", "Washing Date", "Packing Date"]
                output = result[[c for c in final_cols if c in result.columns]].copy()
                output = output.reset_index(drop=True)

                # 4. ทำหน้าสรุป (Summary)
                summary = pd.DataFrame(columns=["Washing Date", "Total Lot"])
                if "Washing Date" in output.columns:
                    # ตัดแถวที่หา Washing Date ไม่เจอออกก่อนทำสรุป
                    valid_output = output.dropna(subset=["Washing Date"])
                    if not valid_output.empty:
                        summary = valid_output.groupby("Washing Date")["Lot"].count().reset_index()
                        summary.columns = ["Washing Date", "Total Lot"]

                # 5. เตรียมไฟล์สำหรับ Download (Excel หลาย Sheet)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    output.to_excel(writer, index=False, sheet_name="Result")
                    if not summary.empty:
                        summary.to_excel(writer, index=False, sheet_name="Summary")
                    else:
                        # สร้าง sheet เปล่าถ้าไม่มีข้อมูล เพื่อกัน error ตอนโหลด
                        pd.DataFrame({"Status": ["No Data Found"]}).to_excel(writer, index=False, sheet_name="Summary")
                
                st.session_state.output = output
                st.session_state.summary = summary
                st.session_state.file = buffer.getvalue()

# =========================
# DISPLAY RESULTS
# =========================
if st.session_state.output is not None:
    st.success("✅ Process สำเร็จ")
    st.subheader("📋 Result")
    st.dataframe(st.session_state.output, use_container_width=True)
    
    if st.session_state.summary is not None and not st.session_state.summary.empty:
        st.subheader("📊 Summary")
        st.dataframe(st.session_state.summary, use_container_width=True)
    else:
        st.warning("⚠️ ไม่พบข้อมูล Washing Date ที่ตรงกับฐานข้อมูล จึงไม่สามารถสร้าง Summary ได้")

    st.download_button(
        label="📥 Download Excel (Result + Summary)",
        data=st.session_state.file,
        file_name="washing_date_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
