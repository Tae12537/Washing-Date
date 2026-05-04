import streamlit as st
import pandas as pd
import io
import re
 
st.title("📊 Washing Date Processor")
 
# =========================
# SESSION STATE
# =========================
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
 
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
# READ
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
        if pd.isna(val):
            break
        val_str = str(val).strip()
        if val_str == "":
            break
        lot_list.append(val_str)
 
    return pd.DataFrame({"Lot": lot_list})
 
def read_file2(file):
    df = read_excel(file)
 
    # หา header row (เหมือนเดิม)
    header_row = None
    for i in range(20):
        row = df.iloc[i].astype(str).str.lower()
        if row.str.contains("runcard").any() and row.str.contains("barcode").any():
            header_row = i
            break
 
    if header_row is None:
        st.error("❌ หา header ไม่เจอ (Runcard / Barcode)")
        return pd.DataFrame()
 
    # ตั้ง header
    df.columns = df.iloc[header_row]
    df = df[header_row + 1:]
 
    df.columns = df.columns.astype(str).str.strip().str.lower()
 
    # 🔥 ของเดิมคุณใช้ตรงนี้ → ไม่แก้
    lot_cols = [c for c in df.columns if "runcard" in str(c).lower()]
    barcode_cols = [c for c in df.columns if "barcode" in str(c).lower()]
 
    if len(lot_cols) == 0 or len(barcode_cols) == 0:
        st.error(f"❌ หา column ไม่เจอ\nColumns ที่มี: {list(df.columns)}")
        return pd.DataFrame()
 
    lot_col = lot_cols[0]
    barcode_col = barcode_cols[0]
 
    # =========================
    # ✅ เพิ่ม Packed Date (ไม่กระทบของเดิม)
    # =========================
    packed_col = None
    for c in df.columns:
        if "packed" in c and "date" in c:
            packed_col = c
            break
 
    if packed_col is None:
        st.error("❌ หา Q4 Packed Date ไม่เจอ")
        return pd.DataFrame()
 
    df_out = df[[lot_col, barcode_col, packed_col]].copy()
    df_out.columns = ["Lot", "Barcode No", "Packed Date"]
 
    df_out = df_out.dropna(subset=["Lot"])
    df_out["Lot"] = df_out["Lot"].astype(str).str.strip()
    df_out["Packed Date"] = pd.to_datetime(df_out["Packed Date"], errors="coerce")
 
    return df_out
 
# =========================
# EXTRACT
# =========================
def extract_ww_day(barcode):
    try:
        s = str(barcode)
        match = re.search('[A-Za-z]', s)
 
        if not match:
            return None, None
 
        start = match.start()
        code = s[start+3:start+6]
 
        if len(code) != 3 or not code.isdigit():
            return None, None
 
        return int(code[:2]), int(code[2])
 
    except:
        return None, None
 
# =========================
# LOAD DATABASE (ใหม่)
# =========================
def load_database():
    df = pd.read_csv("database.txt")
    df["Date"] = pd.to_datetime(df["Date"], format="%d-%b-%Y", errors="coerce")
    return df
 
# =========================
# เลือกวันที่ใกล้สุด
# =========================
def find_best_date(row, date_db):
    if pd.isna(row["WW"]) or pd.isna(row["Day"]) or pd.isna(row["Packed Date"]):
        return None
 
    candidates = date_db[
        (date_db["WW"] == row["WW"]) &
        (date_db["Day"] == row["Day"])
    ].copy()
 
    if candidates.empty:
        return None
 
    candidates["diff"] = (candidates["Date"] - row["Packed Date"]).abs()
    best = candidates.sort_values("diff").iloc[0]
 
    return best["Date"]
 
# =========================
# PROCESS
# =========================
if st.button("🚀 Process"):
 
    if file1 is None or file2 is None:
        st.warning("⚠️ กรุณาอัพโหลดไฟล์ให้ครบ")
    else:
        df1 = read_file1(file1)
        df2 = read_file2(file2)
 
        merged = pd.merge(df1, df2, on="Lot", how="left")
        merged = merged.drop_duplicates(subset=["Lot"])
 
        merged[['WW', 'Day']] = merged['Barcode No'].apply(
            lambda x: pd.Series(extract_ww_day(x))
        )
 
        merged["WW"] = pd.to_numeric(merged["WW"], errors="coerce")
        merged["Day"] = pd.to_numeric(merged["Day"], errors="coerce")
 
        # =========================
        # ✅ ใช้ database.txt แทน
        # =========================
        date_db = load_database()
 
        date_db["WW"] = pd.to_numeric(date_db["WW"], errors="coerce")
        date_db["Day"] = pd.to_numeric(date_db["Day"], errors="coerce")
 
        # =========================
        # 🔥 เลือกวันที่ใกล้ Packed Date
        # =========================
        merged["Washing Date"] = merged.apply(
            lambda row: find_best_date(row, date_db),
            axis=1
        )
 
        output = merged[["Lot", "Barcode No", "WW", "Day", "Washing Date"]].copy()
 
        output = output[output["Lot"].astype(str).str.lower() != "lot/serial"]
        output = output.reset_index(drop=True)
 
        summary = (
            output.groupby("Washing Date")["Lot"]
            .count()
            .reset_index()
            .rename(columns={"Lot": "Total Lot"})
        )
 
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            output.to_excel(writer, index=False, sheet_name="Result")
            summary.to_excel(writer, index=False, sheet_name="Summary")
 
        buffer.seek(0)
 
        st.session_state.output = output
        st.session_state.summary = summary
        st.session_state.file = buffer.getvalue()
 
# =========================
# SHOW
# =========================
if (
    "output" in st.session_state
    and st.session_state.output is not None
):
 
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
 
