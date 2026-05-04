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
    # ดึง Lot จาก Column F (index 5) ตั้งแต่แถวที่ 17 เป็นต้นไป
    data = df.iloc[16:, 5]
    lot_list = []
    for val in data:
        if pd.isna(val) or str(val).strip() == "": break
        lot_list.append(str(val).strip())
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

    lot_col = [c for c in df_data.columns if "runcard" in str(c)][0]
    barcode_col = [c for c in df_data.columns if "barcode" in str(c)][0]
    
    df_out = df_data[[lot_col, barcode_col]].copy()
    df_out.columns = ["Lot", "Barcode No"]
    df_out["Lot"] = df_out["Lot"].astype(str).str.strip()
    return df_out

def extract_logic(barcode):
    """Logic ดึง WW, Day, Year จาก Barcode ตามโครงสร้างที่คุณให้มา"""
    try:
        s = str(barcode)
        match = re.search('[A-Za-z]', s)
        if not match: return None, None, None
        start = match.start()
        # จากตัวอย่าง 760818400AM04336101:
        # start+3:start+5 คือ '43' (WW)
        # start+5 คือ '3' (Day)
        # start+6 คือ '6' (Year 2026)
        ww = int(s[start+3:start+5])
        day = int(s[start+5])
        year_digit = int(s[start+6])
        year = 2020 + year_digit 
        return ww, day, year
    except:
        return None, None, None

# =========================
# MAIN UI
# =========================
st.title("📊 Washing Date Processor")

if "key" not in st.session_state: st.session_state.key = 0

f1 = st.file_uploader("📂 Upload File 1 (Lot)", type=["xls", "xlsx"], key=f"f1_{st.session_state.key}")
f2 = st.file_uploader("📂 Upload File 2 (Barcode)", type=["xls", "xlsx"], key=f"f2_{st.session_state.key}")

if st.button("🚀 Process"):
    if f1 and f2:
        df1 = read_file1(f1)
        df2 = read_file2(f2)

        if not df1.empty and not df2.empty:
            # 1. รวมไฟล์งาน
            merged = pd.merge(df1, df2, on="Lot", how="left")
            
            # 2. ดึงข้อมูล WW, Day, Year จาก Barcode
            info = merged['Barcode No'].apply(lambda x: pd.Series(extract_logic(x)))
            merged[['WW', 'Day', 'Year']] = info

            # 3. โหลด Database ทั้งหมด (2025 และ 2026) มาเชื่อมกัน
            all_dbs = []
            for y in [2025, 2026]:
                path = f"{y}.txt"
                if os.path.exists(path):
                    db = pd.read_csv(path, skipinitialspace=True)
                    db.columns = [c.strip() for c in db.columns]
                    # เปลี่ยนชื่อคอลัมน์ Date เป็น Washing Date
                    if "Date" in db.columns:
                        db = db.rename(columns={"Date": "Washing Date"})
                    # บังคับ Type ให้เป็นตัวเลขทั้งหมด
                    for c in ["Year", "WW", "Day"]:
                        if c in db.columns:
                            db[c] = pd.to_numeric(db[c], errors='coerce')
                    all_dbs.append(db)
            
            if not all_dbs:
                st.error("❌ ไม่พบไฟล์ 2025.txt หรือ 2026.txt ในระบบ")
            else:
                db_final = pd.concat(all_dbs).dropna(subset=["Year", "WW", "Day"])
                db_final[["Year", "WW", "Day"]] = db_final[["Year", "WW", "Day"]].astype(int)

                # 4. บังคับ Type ฝั่งไฟล์งานให้เป็น Int ก่อน Merge
                merged = merged.dropna(subset=["Year", "WW", "Day"])
                merged[["Year", "WW", "Day"]] = merged[["Year", "WW", "Day"]].astype(int)

                # 5. Merge กับ Database
                final_result = pd.merge(merged, db_final, on=["Year", "WW", "Day"], how="left")

                # แสดงผล
                st.success("✅ ดึงข้อมูลสำเร็จ")
                st.dataframe(final_result[["Lot", "Barcode No", "Year", "WW", "Day", "Washing Date"]])

                # 6. ทำหน้าสรุปและดาวน์โหลด
                summary = final_result.dropna(subset=["Washing Date"]).groupby("Washing Date").size().reset_index(name="Total Lot")
                
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    final_result.to_excel(writer, index=False, sheet_name="Result")
                    if not summary.empty:
                        summary.to_excel(writer, index=False, sheet_name="Summary")
                
                st.download_button("📥 Download Excel Report", buf.getvalue(), "washing_report.xlsx")
        else:
            st.error("❌ ข้อมูลในไฟล์ว่างเปล่า หรือหาหัวข้อไม่เจอ")

if st.button("🔄 Reset"):
    st.session_state.key += 1
    st.rerun()
