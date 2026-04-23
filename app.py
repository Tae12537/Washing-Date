import streamlit as st
import pandas as pd
import io
import re

st.title("📊 Washing Date Processor")

# =========================
# Upload
# =========================
file1 = st.file_uploader("📂 Upload File 1 (Lot Shipment)", type=["xlsx", "xls"])
file2 = st.file_uploader("📂 Upload File 2 (Runcard Data)", type=["xlsx", "xls"])

# =========================
# Reset
# =========================
if st.button("🔄 Reset"):
    st.session_state.clear()
    st.rerun()

# =========================
# Read File 1 (Fix Position)
# =========================
def read_file1(file):
    df = pd.read_excel(file, header=None)

    # check format
    if str(df.iloc[15, 5]).strip().lower() != "lot/serial":
        st.error("❌ File 1 format ผิด (F16 ต้องเป็น Lot/Serial)")
        st.stop()

    data = df.iloc[16:, 5]

    df_out = pd.DataFrame()
    df_out["Lot"] = data.astype(str).str.strip()

    df_out = df_out[df_out["Lot"] != ""]
    df_out = df_out.dropna()

    return df_out


# =========================
# Read File 2 (Fix Position)
# =========================
def read_file2(file):
    df = pd.read_excel(file, header=None)

    # check format
    if str(df.iloc[3, 1]).strip().lower() != "runcard no":
        st.error("❌ File 2 format ผิด (B4 ต้องเป็น Runcard No)")
        st.stop()

    if str(df.iloc[3, 8]).strip().lower() != "barcode no":
        st.error("❌ File 2 format ผิด (I4 ต้องเป็น Barcode No)")
        st.stop()

    data = df.iloc[4:, [1, 8]]

    df_out = pd.DataFrame()
    df_out["Lot"] = data.iloc[:, 0].astype(str).str.strip()
    df_out["Barcode No"] = data.iloc[:, 1].astype(str).str.strip()

    df_out = df_out[(df_out["Lot"] != "") & (df_out["Barcode No"] != "")]
    df_out = df_out.dropna()

    return df_out


# =========================
# Extract WW / Day
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
# Process
# =========================
if st.button("🚀 Process"):

    if file1 is None or file2 is None:
        st.warning("❗ กรุณาอัปโหลดไฟล์ให้ครบ")
        st.stop()

    df1 = read_file1(file1)
    df2 = read_file2(file2)

    # merge → เอาเฉพาะ file1
    merged = pd.merge(
        df1,
        df2,
        on="Lot",
        how="left"
    )

    # 1 lot = 1 row
    merged = merged.drop_duplicates(subset=["Lot"])

    # extract WW Day
    merged[['WW', 'Day']] = merged['Barcode No'].apply(
        lambda x: pd.Series(extract_ww_day(x))
    )

    # แปลง type กัน error
    merged["WW"] = pd.to_numeric(merged["WW"], errors="coerce")
    merged["Day"] = pd.to_numeric(merged["Day"], errors="coerce")

    # =========================
    # Date DB (ใส่เต็มได้)
    # =========================
    data = """WW,Day,Date
28,1,03-Jan-2026
28,2,04-Jan-2026
28,3,05-Jan-2026
28,4,06-Jan-2026
28,5,07-Jan-2026
28,6,08-Jan-2026
28,7,09-Jan-2026
29,1,10-Jan-2026
29,2,11-Jan-2026
29,3,12-Jan-2026
29,4,13-Jan-2026
29,5,14-Jan-2026
29,6,15-Jan-2026
29,7,16-Jan-2026
30,1,17-Jan-2026
30,2,18-Jan-2026
30,3,19-Jan-2026
30,4,20-Jan-2026
30,5,21-Jan-2026
30,6,22-Jan-2026
30,7,23-Jan-2026
31,1,24-Jan-2026
31,2,25-Jan-2026
31,3,26-Jan-2026
31,4,27-Jan-2026
31,5,28-Jan-2026
31,6,29-Jan-2026
31,7,30-Jan-2026
32,1,31-Jan-2026
32,2,01-Feb-2026
32,3,02-Feb-2026
32,4,03-Feb-2026
32,5,04-Feb-2026
32,6,05-Feb-2026
32,7,06-Feb-2026
33,1,07-Feb-2026
33,2,08-Feb-2026
33,3,09-Feb-2026
33,4,10-Feb-2026
33,5,11-Feb-2026
33,6,12-Feb-2026
33,7,13-Feb-2026
34,1,14-Feb-2026
34,2,15-Feb-2026
34,3,16-Feb-2026
34,4,17-Feb-2026
34,5,18-Feb-2026
34,6,19-Feb-2026
34,7,20-Feb-2026
35,1,21-Feb-2026
35,2,22-Feb-2026
35,3,23-Feb-2026
35,4,24-Feb-2026
35,5,25-Feb-2026
35,6,26-Feb-2026
35,7,27-Feb-2026
36,1,28-Feb-2026
36,2,01-Mar-2026
36,3,02-Mar-2026
36,4,03-Mar-2026
36,5,04-Mar-2026
36,6,05-Mar-2026
36,7,06-Mar-2026
37,1,07-Mar-2026
37,2,08-Mar-2026
37,3,09-Mar-2026
37,4,10-Mar-2026
37,5,11-Mar-2026
37,6,12-Mar-2026
37,7,13-Mar-2026
38,1,14-Mar-2026
38,2,15-Mar-2026
38,3,16-Mar-2026
38,4,17-Mar-2026
38,5,18-Mar-2026
38,6,19-Mar-2026
38,7,20-Mar-2026
39,1,21-Mar-2026
39,2,22-Mar-2026
39,3,23-Mar-2026
39,4,24-Mar-2026
39,5,25-Mar-2026
39,6,26-Mar-2026
39,7,27-Mar-2026
40,1,28-Mar-2026
40,2,29-Mar-2026
40,3,30-Mar-2026
40,4,31-Mar-2026
40,5,01-Apr-2026
40,6,02-Apr-2026
40,7,03-Apr-2026
41,1,04-Apr-2026
41,2,05-Apr-2026
41,3,06-Apr-2026
41,4,07-Apr-2026
41,5,08-Apr-2026
41,6,09-Apr-2026
41,7,10-Apr-2026
42,1,11-Apr-2026
42,2,12-Apr-2026
42,3,13-Apr-2026
42,4,14-Apr-2026
42,5,15-Apr-2026
42,6,16-Apr-2026
42,7,17-Apr-2026
43,1,18-Apr-2026
43,2,19-Apr-2026
43,3,20-Apr-2026
43,4,21-Apr-2026
43,5,22-Apr-2026
43,6,23-Apr-2026
43,7,24-Apr-2026
44,1,25-Apr-2026
44,2,26-Apr-2026
44,3,27-Apr-2026
44,4,28-Apr-2026
44,5,29-Apr-2026
44,6,30-Apr-2026
44,7,01-May-2026
45,1,02-May-2026
45,2,03-May-2026
45,3,04-May-2026
45,4,05-May-2026
45,5,06-May-2026
45,6,07-May-2026
45,7,08-May-2026
46,1,09-May-2026
46,2,10-May-2026
46,3,11-May-2026
46,4,12-May-2026
46,5,13-May-2026
46,6,14-May-2026
46,7,15-May-2026
47,1,16-May-2026
47,2,17-May-2026
47,3,18-May-2026
47,4,19-May-2026
47,5,20-May-2026
47,6,21-May-2026
47,7,22-May-2026
48,1,23-May-2026
48,2,24-May-2026
48,3,25-May-2026
48,4,26-May-2026
48,5,27-May-2026
48,6,28-May-2026
48,7,29-May-2026
49,1,30-May-2026
49,2,31-May-2026
49,3,01-Jun-2026
49,4,02-Jun-2026
49,5,03-Jun-2026
49,6,04-Jun-2026
49,7,05-Jun-2026
50,1,06-Jun-2026
50,2,07-Jun-2026
50,3,08-Jun-2026
50,4,09-Jun-2026
50,5,10-Jun-2026
50,6,11-Jun-2026
50,7,12-Jun-2026
51,1,13-Jun-2026
51,2,14-Jun-2026
51,3,15-Jun-2026
51,4,16-Jun-2026
51,5,17-Jun-2026
51,6,18-Jun-2026
51,7,19-Jun-2026
52,1,20-Jun-2026
52,2,21-Jun-2026
52,3,22-Jun-2026
52,4,23-Jun-2026
52,5,24-Jun-2026
52,6,25-Jun-2026
52,7,26-Jun-2026
53,1,27-Jun-2026
53,2,28-Jun-2026
53,3,29-Jun-2026
53,4,30-Jun-2026
53,5,01-Jul-2026
53,6,02-Jul-2026
53,7,03-Jul-2026
1,1,04-Jul-2026
1,2,05-Jul-2026
1,3,06-Jul-2026
1,4,07-Jul-2026
1,5,08-Jul-2026
1,6,09-Jul-2026
1,7,10-Jul-2026
2,1,11-Jul-2026
2,2,12-Jul-2026
2,3,13-Jul-2026
2,4,14-Jul-2026
2,5,15-Jul-2026
2,6,16-Jul-2026
2,7,17-Jul-2026
3,1,18-Jul-2026
3,2,19-Jul-2026
3,3,20-Jul-2026
3,4,21-Jul-2026
3,5,22-Jul-2026
3,6,23-Jul-2026
3,7,24-Jul-2026
4,1,25-Jul-2026
4,2,26-Jul-2026
4,3,27-Jul-2026
4,4,28-Jul-2026
4,5,29-Jul-2026
4,6,30-Jul-2026
4,7,31-Jul-2026
5,1,01-Aug-2026
5,2,02-Aug-2026
5,3,03-Aug-2026
5,4,04-Aug-2026
5,5,05-Aug-2026
5,6,06-Aug-2026
5,7,07-Aug-2026
6,1,08-Aug-2026
6,2,09-Aug-2026
6,3,10-Aug-2026
6,4,11-Aug-2026
6,5,12-Aug-2026
6,6,13-Aug-2026
6,7,14-Aug-2026
7,1,15-Aug-2026
7,2,16-Aug-2026
7,3,17-Aug-2026
7,4,18-Aug-2026
7,5,19-Aug-2026
7,6,20-Aug-2026
7,7,21-Aug-2026
8,1,22-Aug-2026
8,2,23-Aug-2026
8,3,24-Aug-2026
8,4,25-Aug-2026
8,5,26-Aug-2026
8,6,27-Aug-2026
8,7,28-Aug-2026
9,1,29-Aug-2026
9,2,30-Aug-2026
9,3,31-Aug-2026
9,4,01-Sep-2026
9,5,02-Sep-2026
9,6,03-Sep-2026
9,7,04-Sep-2026
10,1,05-Sep-2026
10,2,06-Sep-2026
10,3,07-Sep-2026
10,4,08-Sep-2026
10,5,09-Sep-2026
10,6,10-Sep-2026
10,7,11-Sep-2026
11,1,12-Sep-2026
11,2,13-Sep-2026
11,3,14-Sep-2026
11,4,15-Sep-2026
11,5,16-Sep-2026
11,6,17-Sep-2026
11,7,18-Sep-2026
12,1,19-Sep-2026
12,2,20-Sep-2026
12,3,21-Sep-2026
12,4,22-Sep-2026
12,5,23-Sep-2026
12,6,24-Sep-2026
12,7,25-Sep-2026
13,1,26-Sep-2026
13,2,27-Sep-2026
13,3,28-Sep-2026
13,4,29-Sep-2026
13,5,30-Sep-2026
13,6,01-Oct-2026
13,7,02-Oct-2026
14,1,03-Oct-2026
14,2,04-Oct-2026
14,3,05-Oct-2026
14,4,06-Oct-2026
14,5,07-Oct-2026
14,6,08-Oct-2026
14,7,09-Oct-2026
15,1,10-Oct-2026
15,2,11-Oct-2026
15,3,12-Oct-2026
15,4,13-Oct-2026
15,5,14-Oct-2026
15,6,15-Oct-2026
15,7,16-Oct-2026
16,1,17-Oct-2026
16,2,18-Oct-2026
16,3,19-Oct-2026
16,4,20-Oct-2026
16,5,21-Oct-2026
16,6,22-Oct-2026
16,7,23-Oct-2026
17,1,24-Oct-2026
17,2,25-Oct-2026
17,3,26-Oct-2026
17,4,27-Oct-2026
17,5,28-Oct-2026
17,6,29-Oct-2026
17,7,30-Oct-2026
18,1,31-Oct-2026
18,2,01-Nov-2026
18,3,02-Nov-2026
18,4,03-Nov-2026
18,5,04-Nov-2026
18,6,05-Nov-2026
18,7,06-Nov-2026
19,1,07-Nov-2026
19,2,08-Nov-2026
19,3,09-Nov-2026
19,4,10-Nov-2026
19,5,11-Nov-2026
19,6,12-Nov-2026
19,7,13-Nov-2026
20,1,14-Nov-2026
20,2,15-Nov-2026
20,3,16-Nov-2026
20,4,17-Nov-2026
20,5,18-Nov-2026
20,6,19-Nov-2026
20,7,20-Nov-2026
21,1,21-Nov-2026
21,2,22-Nov-2026
21,3,23-Nov-2026
21,4,24-Nov-2026
21,5,25-Nov-2026
21,6,26-Nov-2026
21,7,27-Nov-2026
22,1,28-Nov-2026
22,2,29-Nov-2026
22,3,30-Nov-2026
22,4,01-Dec-2026
22,5,02-Dec-2026
22,6,03-Dec-2026
22,7,04-Dec-2026
23,1,05-Dec-2026
23,2,06-Dec-2026
23,3,07-Dec-2026
23,4,08-Dec-2026
23,5,09-Dec-2026
23,6,10-Dec-2026
23,7,11-Dec-2026
24,1,12-Dec-2026
24,2,13-Dec-2026
24,3,14-Dec-2026
24,4,15-Dec-2026
24,5,16-Dec-2026
24,6,17-Dec-2026
24,7,18-Dec-2026
25,1,19-Dec-2026
25,2,20-Dec-2026
25,3,21-Dec-2026
25,4,22-Dec-2026
25,5,23-Dec-2026
25,6,24-Dec-2026
25,7,25-Dec-2026
26,1,26-Dec-2026
26,2,27-Dec-2026
26,3,28-Dec-2026
26,4,29-Dec-2026
26,5,30-Dec-2026
26,6,31-Dec-2026

"""

    date_db = pd.read_csv(io.StringIO(data))
    date_db["WW"] = pd.to_numeric(date_db["WW"])
    date_db["Day"] = pd.to_numeric(date_db["Day"])

    # merge date
    result = pd.merge(merged, date_db, on=["WW", "Day"], how="left")

    output = result[["Lot", "Barcode No", "Date"]]
    output = output.rename(columns={"Date": "Washing Date"})

    # =========================
    # Summary
    # =========================
    summary = (
        output.groupby("Washing Date")["Lot"]
        .count()
        .reset_index()
        .rename(columns={"Lot": "Total Lot"})
    )

    # =========================
    # Export Excel
    # =========================
    output_file = io.BytesIO()

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        output.to_excel(writer, index=False, sheet_name="Result")

        start_row = len(output) + 3
        summary.to_excel(writer, index=False, sheet_name="Result", startrow=start_row)

    # =========================
    # Show Result
    # =========================
    st.success("✅ เสร็จแล้ว")
    st.dataframe(output)

    st.download_button(
        "📥 Download Excel",
        data=output_file.getvalue(),
        file_name="washing_date_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
