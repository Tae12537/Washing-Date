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
# UPLOAD SECTION
# =========================
file1 = st.file_uploader(
    "📂 Upload File 1 (Lot/Serial)",
    type=["xls", "xlsx", "csv"],
    key=f"file1_{st.session_state.uploader_key}"
)

file2 = st.file_uploader(
    "📂 Upload File 2 (Runcard / Barcode)",
    type=["xls", "xlsx", "csv"],
    key=f"file2_{st.session_state.uploader_key}"
)

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
    df_data = df[header_row + 1:].copy()
    df_data.columns = df_data.columns.astype(str).str.strip().str.lower()

    lot_cols = [c for c in df_data.columns if "runcard" in str(c).lower()]
    barcode_cols = [c for c in df_data.columns if "barcode" in str(c).lower()]
    # หาคอลัมน์ Q4 หรือ Packing Date
    packing_cols = [c for c in df_data.columns if "packing" in str(c).lower() or "q4" == str(c).lower()]

    if not lot_cols or not barcode_cols:
        st.error("❌ ไม่พบคอลัมน์ Runcard หรือ Barcode")
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
        df_out["Packing Date"] = pd.to_datetime(df_out["Packing Date"], errors='coerce')
        
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
# BUTTONS (PROCESS & RESET SIDE BY SIDE)
# =========================
col1, col2, _ = st.columns([1, 1, 4])

with col1:
    btn_process = st.button("🚀 Process", use_container_width=True)

with col2:
    if st.button("🔄 Reset", use_container_width=True):
        for key in ["output", "summary", "file"]:
            st.session_state[key] = None
        st.session_state.uploader_key += 1
        st.rerun()

# =========================
# MAIN LOGIC
# =========================
if btn_process:
    if file1 is None or file2 is None:
        st.warning("⚠️ กรุณาอัพโหลดไฟล์ให้ครบ")
    else:
        df1 = read_file1(file1)
        df2 = read_file2(file2)

        if not df2.empty:
            # 1. ดึงปีจาก Packing Date (เอาค่าที่มากที่สุดมาเป็นตัวตั้งต้นโหลด DB)
            if "Packing Date" in df2.columns and not df2["Packing Date"].isnull().all():
                detected_year = df2["Packing Date"].dt.year.max()
            else:
                detected_year = 2026 # Default
            
            db_filename = f"{int(detected_year)}.txt"
            
            if not os.path.exists(db_filename):
                st.error(f"❌ ไม่พบไฟล์ฐานข้อมูล `{db_filename}`")
            else:
                # 2. โหลด Database จากไฟล์ .txt (Year,WW,Day,Date)
                date_db = pd.read_csv(db_filename)
                date_db.columns = date_db.columns.str.strip()
                date_db["Year"] = pd.to_numeric(date_db["Year"], errors="coerce")
                date_db["WW"] = pd.to_numeric(date_db["WW"], errors="coerce")
                date_db["Day"] = pd.to_numeric(date_db["Day"], errors="coerce")

                # 3. เตรียมข้อมูลจากไฟล์ที่อัพโหลด
                merged = pd.merge(df1, df2, on="Lot", how="left").drop_duplicates(subset=["Lot"])
                merged[['WW', 'Day']] = merged['Barcode No'].apply(lambda x: pd.Series(extract_ww_day(x)))
                
                merged["WW"] = pd.to_numeric(merged["WW"], errors="coerce")
                merged["Day"] = pd.to_numeric(merged["Day"], errors="coerce")
                
                # ดึง Year จาก Packing Date ในแต่ละแถวมาใช้ merge
                if "Packing Date" in merged.columns:
                    merged["Year"] = merged["Packing Date"].dt.year
                else:
                    merged["Year"] = detected_year

                # 4. Merge กับ DB โดยใช้ Year, WW, Day
                result = pd.merge(merged, date_db, on=["Year", "WW", "Day"], how="left")

                # 5. สรุปผล
                output = result[["Lot", "Barcode No", "Year", "WW", "Day", "Date", "Packing Date"]].copy()
                output = output.rename(columns={"Date": "Washing Date"})
                output = output[output["Lot"].astype(str).str.lower() != "lot/serial"]
                output = output.reset_index(drop=True)

                summary = (
                    output.groupby("Washing Date")["Lot"]
                    .count()
                    .reset_index()
                    .rename(columns={"Lot": "Total Lot"})
                )

                # Export to Excel
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    output.to_excel(writer, index=False, sheet_name="Result")
                    summary.to_excel(writer, index=False, sheet_name="Summary")
                
                st.session_state.output = output
                st.session_state.summary = summary
                st.session_state.file = buffer.getvalue()

# =========================
# DISPLAY RESULT
# =========================
if st.session_state.output is not None:
    st.success("✅ Process สำเร็จ")
    st.subheader("📋 Result")
    st.dataframe(st.session_state.output, use_container_width=True)

    st.subheader("📊 Summary")
    st.dataframe(st.session_state.summary, use_container_width=True)

    st.download_button(
        label="📥 Download Excel",
        data=st.session_state.file,
        file_name="washing_date_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
