import pandas as pd
import numpy as np
from datetime import datetime
import os

# ==========================================
# 1. โหลด DATABASE (WASHING CALENDAR)
# ==========================================
def load_database(db_path='database.txt'):
    if not os.path.exists(db_path):
        print(f"Error: ไม่พบไฟล์ {db_path}")
        return None
    try:
        # อ่าน database ที่มี Year อยู่คอลัมน์แรก
        df_db = pd.read_csv(db_path)
        # แปลงวันที่ให้เป็น datetime format
        df_db['Date'] = pd.to_datetime(df_db['Date'], format='%d-%b-%Y')
        return df_db
    except Exception as e:
        print(f"Error ในการโหลด Database: {e}")
        return None

# ==========================================
# 2. LOGIC การหา WASHING DATE ที่ถูกต้อง
# ==========================================
def get_washing_date(lot_val, packed_date_val, df_db):
    try:
        # 1. ตรวจสอบค่าว่าง
        if pd.isna(lot_val) or pd.isna(packed_date_val):
            return None
        
        # 2. แกะ WW และ Day จาก Lot (เช่น 127501 -> WW=27, Day=5)
        lot_str = str(lot_val).strip()
        if len(lot_str) < 4: return "Lot Invalid"
        
        target_ww = int(lot_str[1:3])
        target_day = int(lot_str[3:4])

        # 3. จัดการ Packed Date (Column Q) 
        # รองรับทั้งแบบ String และ Datetime object
        if isinstance(packed_date_val, str):
            # ตัดเอาแค่ส่วนวันที่ '27/11/2025' จาก '27/11/2025 23:51:07'
            date_only_str = packed_date_val.split(' ')[0]
            packed_dt = datetime.strptime(date_only_str, '%d/%m/%Y')
        else:
            # ถ้ามาเป็น datetime อยู่แล้ว ให้ตัดเวลาทิ้งเพื่อเทียบแค่วันที่
            packed_dt = packed_date_val.replace(hour=0, minute=0, second=0, microsecond=0)

        # 4. ค้นหาใน Database
        # เลือกแถวที่ WW และ Day ตรงกัน
        matches = df_db[(df_db['WW'] == target_ww) & (df_db['Day'] == target_day)].copy()

        if matches.empty:
            return "No WW/Day Match"

        # 5. กรองเอาเฉพาะวันที่ "ก่อนหน้า" (Before) วันที่แพ็ค (Packed Date)
        valid_dates = matches[matches['Date'] < packed_dt]

        if not valid_dates.empty:
            # เรียงลำดับจากใหม่ไปเก่า (Descending) แล้วเลือกอันที่ใกล้ที่สุด
            best_match = valid_dates.sort_values(by='Date', ascending=False).iloc[0]
            return best_match['Date'].strftime('%d-%b-%Y')
        else:
            # กรณีไม่มีวันไหนใน DB ที่เก่ากว่าวันแพ็คเลย
            return "Check Year (DB > Packed)"

    except Exception as e:
        return f"Error: {str(e)}"

# ==========================================
# 3. ส่วนประมวลผลไฟล์หลัก (MAIN PROCESS)
# ==========================================
def main():
    # --- ตั้งค่าชื่อไฟล์ตรงนี้ ---
    input_file = 'your_work_file.xlsx'  # ชื่อไฟล์ Excel ของคุณ
    output_file = 'Result_WashingDate.xlsx'
    db_file = 'database.txt'
    
    print("--- เริ่มการทำงาน ---")
    
    # โหลด Database
    df_db = load_database(db_file)
    if df_db is None: return

    # อ่านไฟล์ Excel
    try:
        # อ่านไฟล์โดยระบุว่า Column Q (Index 16) คือ Packed Date
        df = pd.read_excel(input_file)
        print(f"อ่านไฟล์ {input_file} สำเร็จ: {len(df)} แถว")
    except Exception as e:
        print(f"ไม่สามารถอ่านไฟล์ Excel ได้: {e}")
        return

    # ระบุชื่อคอลัมน์ (แก้ชื่อให้ตรงกับ Excel ของคุณ)
    # สมมติ Column Q ของคุณชื่อ 'Packed Date' และ Column Lot ชื่อ 'Lot'
    # หากไม่ทราบชื่อคอลัมน์ สามารถใช้ Index แทนได้ เช่น df.columns[16]
    col_lot = 'Lot' 
    col_packed = 'Packed Date' 

    if col_lot not in df.columns or col_packed not in df.columns:
        print(f"Error: ไม่พบชื่อคอลัมน์ '{col_lot}' หรือ '{col_packed}' ในไฟล์")
        print(f"คอลัมน์ที่มีอยู่คือ: {list(df.columns)}")
        return

    # --- ขั้นตอนการรัน Logic ---
    print("กำลังคำนวณ Washing Date...")
    df['Washing_Date_Result'] = df.apply(
        lambda row: get_washing_date(row[col_lot], row[col_packed], df_db), 
        axis=1
    )

    # --- การบันทึกผล ---
    try:
        df.to_excel(output_file, index=False)
        print("-" * 30)
        print(f"เสร็จสมบูรณ์!")
        print(f"บันทึกผลไปที่: {output_file}")
        print("-" * 30)
    except Exception as e:
        print(f"Error ในการบันทึกไฟล์: {e}")

if __name__ == "__main__":
    main()
