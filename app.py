import streamlit as st
import pandas as pd
import io
import re
import os

st.set_page_config(page_title="Washing Date Processor", layout="wide")

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
    if header_row is None: return pd.DataFrame()

    df.columns = df.iloc[header_row]
    df_data = df[header_row + 1:].copy()
    df_data.columns = df_data.columns.astype(str).str.strip().str.lower()

    # หาคอลัมน์ Lot และ Barcode
    lot_col = [c for c in df_data.columns if "runcard" in str(c)][0]
    barcode_col = [c for c in df_data.columns if "barcode" in str(c)][0]
    packing_col = [c for c in df_data.columns if "packing" in str(c) or str(c) == "q4"]

    cols = [lot_col, barcode_col]
    if packing_col: cols.append(packing_col[0])

    df_out = df_data[cols].copy()
    df_out.columns = ["Lot", "Barcode No"] + (["Packing Date"] if packing_col else [])
    df_out = df_out.dropna(subset=["Lot"])
    df_out["Lot"] = df_out["Lot"].astype(str).str.strip()
    return df_out

def extract_info(barcode):
    """ดึง WW, Day และ Year จาก Barcode โดยตรง"""
    try:
        s = str(barcode)
        match = re.search('[A-Za-z]', s)
        if not match: return None, None, None
        start = match.start()
        # WW = ตำแหน่งตัวอักษรตัวสุดท้าย + 4 และ 5
        # Day = ตำแหน่งตัวอักษรตัวสุดท้าย + 6
        # Year = ตำแหน่งตัวอักษรตัวสุดท้าย + 7
        ww = int(s[start+3:start+5])
        day = int(s[start+5])
        year_digit = int(s[start+6])
        full_year = 2020 + year_digit # เช่น 6 -> 2026
        return ww, day, full_year
    except:
        return None, None, None

# =========================
# UI
# =========================
st.title("📊 Washing Date Processor (Fixed)")

if "uploader_key" not in st.session_state: st.session_state.uploader_key = 0

file1 = st.file_uploader("📂 Upload File 1", type=["xls", "xlsx", "csv"], key=f"f1_{st.session_state.uploader_key}")
file2 = st.file_uploader("📂 Upload File 2", type=["xls", "xlsx", "csv"], key=f"f2_{st.session_state.uploader_key}")

if st.button("🚀 Process"):
    if file1 and file2:
        df1 = read_file1(file1)
        df2 = read_file2(file2)

        if not df1.empty and not df2.empty:
            # 1. รวมไฟล์ 1 และ 2
            merged = pd.merge(df1, df2, on="Lot", how="left").drop_duplicates(subset=["Lot"])

            # 2. ดึงข้อมูลจาก Barcode
            extracted = merged['Barcode No'].apply(lambda x: pd.Series(extract_info(x)))
            merged[['WW', 'Day', 'Year']] = extracted

            # 3. โหลด Database (โหลดทั้ง 2025 และ 2026 มาต่อกันกันเหนียว)
            db_list = []
            for y in [2025, 2026]:
                path = f"{y}.txt"
                if os.path.exists(path):
                    temp_db = pd.read_csv(path, skipinitialspace=True)
                    temp_db.columns = [c.strip() for c in temp_db.columns]
                    # เปลี่ยนชื่อคอลัมน์ให้ตรง
                    if "Date" in temp_db.columns: temp_db = temp_db.rename(columns={"Date": "Washing Date"})
                    # บังคับ Type
                    for c in ["Year", "WW", "Day"]:
                        if c in temp_db.columns:
                            temp_db[c] = pd.to_numeric(temp_db[c], errors='coerce')
                    db_list.append(temp_db)
            
            if not db_list:
                st.error("❌ ไม่พบไฟล์ฐานข้อมูล 2025.txt หรือ 2026.txt")
            else:
                full_db = pd.concat(db_list).dropna(subset=["Year", "WW", "Day"])
                full_db[["Year", "WW", "Day"]] = full_db[["Year", "WW", "Day"]].astype(int)

                # 4. Merge กับ Database
                # บังคับ Type ฝั่งไฟล์งานก่อน Merge
                merged[["Year", "WW", "Day"]] = merged[["Year", "WW", "Day"]].fillna(0).astype(int)
                
                final = pd.merge(merged, full_db, on=["Year", "WW", "Day"], how="left")

                # 5. แสดงผล
                show_cols = ["Lot", "Barcode No", "Year", "WW", "Day", "Washing Date"]
                output = final[[c for c in show_cols if c in final.columns]].copy()
                
                st.success("✅ ประมวลผลสำเร็จ")
                st.dataframe(output, use_container_width=True)

                # 6. Summary & Download
                valid_data = output.dropna(subset=["Washing Date"])
                if not valid_data.empty:
                    summary = valid_data.groupby("Washing Date").size().reset_index(name="Total Lot")
                    st.subheader("📊 Summary")
                    st.dataframe(summary)

                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine="openpyxl") as wr:
                        output.to_excel(wr, index=False, sheet_name="Result")
                        summary.to_excel(wr, index=False, sheet_name="Summary")
                    st.download_button("📥 Download Excel", buf.getvalue(), "washing_report.xlsx")
                else:
                    st.warning("⚠️ Washing Date ไม่ขึ้น! กรุณาเช็คว่า Year, WW, Day ในไฟล์งาน ตรงกับในไฟล์ .txt หรือไม่")
                    # Debug: แสดงค่าที่ดึงได้
                    st.write("Debug ข้อมูลที่ดึงได้จากไฟล์งาน (3 แถวแรก):")
                    st.write(merged[["Year", "WW", "Day"]].head(3))

if st.button("🔄 Reset"):
    st.session_state.uploader_key += 1
    st.rerun()
