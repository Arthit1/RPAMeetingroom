import sys
import math
import re
import os
import json
from pathlib import Path
import pandas as pd
import calendar

DOWNLOAD_DIR = Path("downloads")
OUTPUT_DIR = Path("output")
OUTPUT_FILE = OUTPUT_DIR / "merged_report.xlsx"
DEBUG_RAW_FILE = OUTPUT_DIR / "debug_raw_concat.xlsx"

OUTPUT_DIR.mkdir(exist_ok=True)

# ===================================================
# 1. ระบบรับค่าดักจับจาก Pipeline (Open Days & New Files)
# ===================================================
USER_NUM_DAYS = 19 
NEW_FILES_LIST = set() # เก็บรายชื่อไฟล์ที่เพิ่งโหลดมาใหม่ในรอบนี้

# แกะอาร์กิวเมนต์ที่ส่งมาจาก run_pipeline.py
if len(sys.argv) >= 2:
    try:
        USER_NUM_DAYS = int(sys.argv[-2]) # จำนวนวันใหม่จาก UI หน้าจอ
    except (ValueError, IndexError):
        try:
            USER_NUM_DAYS = int(sys.argv[-1])
        except ValueError:
            USER_NUM_DAYS = 19

if len(sys.argv) >= 3:
    try:
        # โหลดรายชื่อไฟล์ใหม่ที่ส่งผ่านมาในรูปแบบ JSON string อัตโนมัติ
        NEW_FILES_LIST = set(json.loads(sys.argv[-1]))
    except Exception:
        NEW_FILES_LIST = set()

HOURS_PER_DAY = 8.5

# ⚡ โครงสร้างลำดับคอลัมน์มาตรฐาน
FINAL_COLUMNS = [
    "No", "Building", 
    "From Year", "From Month", "From Day", 
    "To Year", "To Month", "To Day", 
    "Room", "Number of Bookings", "Number of Meet Now", "Number of Advanced Bookings", 
    "Total Duration", "Total Duration (Hours)", "Avg Duration", "Min Duration", 
    "Max Duration", "Open Days", "Available Hours", "Utilization (%)", 
    "Source File", "Source Sheet", "File Date"
]

EXPECTED_OUTPUT_COLUMNS = [
    "No", "Room", "Number of Bookings", "Number of Meet Now", 
    "Number of Advanced Bookings", "Total Duration", "Avg Duration", 
    "Min Duration", "Max Duration"
]

def extract_date_from_filename(filename):
    match = re.search(r"\((\d{2})(\d{2})(\d{4})", filename)
    if match:
        day, month, year = match.groups()
        return pd.to_datetime(f"{year}-{month}-{day}")
    return None

def parse_room_key(room):
    if pd.isna(room): return (9, 999999, 999999, "")
    text = str(room).strip().replace(".", "")
    upper_text = text.upper()
    m = re.match(r"^(\d+)-(\d+)$", upper_text)
    if m: return (0, int(m.group(1)), int(m.group(2)), upper_text)
    m = re.match(r"^([A-Z]+)(\d+)-(\d+)$", upper_text)
    if m: return (1, int(m.group(2)), int(m.group(3)), m.group(1))
    nums = re.findall(r"\d+", upper_text)
    if nums:
        first = int(nums[0])
        second = int(nums[1]) if len(nums) > 1 else 0
        return (2, first, second, upper_text)
    return (3, 999999, 999999, upper_text)

def parse_duration_to_hours(duration_str):
    if pd.isna(duration_str) or str(duration_str).strip() == "" or str(duration_str).strip() == "0":
        return 0.0
    try:
        duration_str = str(duration_str).strip()
        day_match = re.match(r"(?:(\d+)\s+days?,\s*)?([\d:]+)", duration_str)
        days_offset = 0
        if day_match and day_match.group(1):
            days_offset = int(day_match.group(1)) * 24
            duration_str = day_match.group(2)

        parts = duration_str.split(':')
        if len(parts) >= 2:
            hours = int(parts[0]) + days_offset
            minutes = int(parts[1])
            return round(hours + (minutes / 60.0), 1)
        return round(float(duration_str), 1)
    except:
        return 0.0

def extract_building_name(filename):
    try:
        name = filename.split('(')[0]
        if '_report_' in name:
            name = name.split('_report_')[0]
        else:
            name = name.rsplit('_', 1)[0]
        return name.replace('_', ' ').strip()
    except:
        return "Unknown Building"

def looks_like_header_row(row_values):
    row_text = [str(v).strip().lower() for v in row_values]
    joined = " | ".join(row_text)
    keywords = ["room", "number of bookings", "total duration", "avg duration"]
    return sum(1 for k in keywords if k in joined) >= 3

def normalize_sheet(df, source_file, source_sheet):
    df = df.dropna(how="all").copy().reset_index(drop=True)
    start_idx = 0
    for i in range(min(10, len(df))):
        if looks_like_header_row(df.iloc[i].tolist()):
            start_idx = i + 1
            break

    if start_idx > 0:
        df = df.iloc[start_idx:].copy().reset_index(drop=True)

    if df.shape[1] < 9:
        return None

    df = df.iloc[:, :9].copy()
    df.columns = EXPECTED_OUTPUT_COLUMNS

    header_mask = df.apply(lambda r: looks_like_header_row(r.tolist()), axis=1)
    df = df[~header_mask].copy()

    df["Room"] = df["Room"].astype(str).str.strip()
    
    invalid_keywords = ['undefined', 'undifine', 'total', 'nan', 'summary', '']
    df = df[df["Room"].notna() & (df["Room"] != "")]
    df = df[~df["Room"].str.lower().isin(invalid_keywords)].copy()

    df["Source File"] = source_file
    df["Source Sheet"] = source_sheet
    return df

# 💡 ฟังก์ชันหัวใจสำคัญในการคำนวณและดักจับเงื่อนไข Open Days พิเศษตามโจทย์ของพี่
def calculate_open_days(row, input_days_from_ui, new_files_set):
    building = str(row.get('Building', '')).strip()
    room = str(row.get('Room', '')).strip()
    source_file = str(row.get('Source File', '')).strip()
    
    # 📌 เงื่อนไข 1: Building = THE TARA และ Room ขึ้นต้นด้วย 22- -> บังคับเป็น 22 วัน
    if building == "THE TARA" and room.startswith("22-"):
        return 22
        
    # 📌 เงื่อนไข 2: ตารางจับคู่ห้องชั้นใต้ดินเจาะจง (.B / .B1 / B)
    b1_rooms_mapping = {
        ".B1-01": 20, ".B1-02": 20, ".B1-03": 20, ".B1-04": 21,
        ".B1-05": 22, ".B1-06": 20, ".B1-07": 20, ".B1-08": 21
    }
    
    clean_room = room.replace(".", "").upper()
    for b1_key, days_val in b1_rooms_mapping.items():
        clean_key = b1_key.replace(".", "").upper()
        if clean_key == clean_room or b1_key in room.upper():
            return days_val

    # 📌 เงื่อนไข 3: เช็กจากสถานะไฟล์เก่า-ใหม่
    # ถ้าเป็นไฟล์ที่เพิ่งโหลดมาใหม่รอบนี้ในโฟลเดอร์ดาวน์โหลด -> ใช้ค่าที่กรอกมาใหม่จาก UI
    if source_file in new_files_set:
        return input_days_from_ui
    
    # ถ้าเป็นไฟล์เก่าที่เคยมีอยู่แล้ว -> ดึงค่าเดิมของมันมาใช้ (ถ้าดึงไม่ได้ให้เปลี่ยนเป็นค่าเริ่มต้น 19)
    existing_days = row.get('Open Days', None)
    if pd.notna(existing_days):
        try:
            return int(float(existing_days))
        except ValueError:
            pass
            
    return 19

# ===================================================
# 2. กวาดและประมวลผลไฟล์ทั้งหมดใน downloads
# ===================================================
files = sorted([f for f in DOWNLOAD_DIR.glob("*.xlsx")])

if not files:
    print(f"❌ ไม่พบไฟล์ Excel ใด ๆ เลยในโฟลเดอร์ {DOWNLOAD_DIR.absolute()}")
    raise SystemExit

all_data = []
debug_summary = []
detected_date_objects = [] 

for file in files:
    try:
        file_date = extract_date_from_filename(file.name)
        if file_date:
            detected_date_objects.append(file_date)

        xl = pd.ExcelFile(file)
        building_name = extract_building_name(file.name)
        for sheet_name in xl.sheet_names:
            raw_df = pd.read_excel(file, sheet_name=sheet_name, header=None)
            
            # ตรวจสอบเบื้องต้นเพื่อดึงค่า Open Days ของเดิมที่มีอยู่แล้วในไฟล์เก่า (ถ้ามี)
            existing_days_val = None
            if raw_df.shape[1] >= 18:
                for idx in range(min(15, len(raw_df))):
                    val = raw_df.iloc[idx, 17]
                    if pd.notna(val) and str(val).isdigit():
                        existing_days_val = int(val)
                        break

            normalized_df = normalize_sheet(raw_df, file.name, sheet_name)

            if normalized_df is not None and not normalized_df.empty:
                normalized_df["Building"] = building_name
                normalized_df["File Date"] = file_date  
                normalized_df["Open Days"] = existing_days_val
                all_data.append(normalized_df)
                debug_summary.append({"file": file.name, "sheet": sheet_name, "rows": len(normalized_df), "status": "OK"})
    except Exception as e:
        debug_summary.append({"file": file.name, "sheet": "Unknown", "rows": 0, "status": f"Error: {e}"})

if not all_data:
    print("❌ ไม่มีข้อมูลห้องประชุมที่ใช้งานได้ให้ประมวลผล")
    raise SystemExit

merged = pd.concat(all_data, ignore_index=True)

invalid_global = ['undefined', 'undifine', 'total', 'summary', '']
merged = merged[~merged["Room"].astype(str).str.strip().str.lower().isin(invalid_global)].copy()

merged["_sort"] = merged["Room"].apply(parse_room_key)
merged = merged.sort_values("_sort", kind="stable").drop(columns="_sort").reset_index(drop=True)

merged["Total Duration (Hours)"] = merged["Total Duration"].apply(parse_duration_to_hours)
merged["No"] = range(1, len(merged) + 1)

# 🔥 คำนวณคอลัมน์ Open Days ตามตรรกะเงื่อนไขพิเศษทั้งหมดแบบแถวต่อแถว
merged["Open Days"] = merged.apply(lambda r: calculate_open_days(r, USER_NUM_DAYS, NEW_FILES_LIST), axis=1)

# คำนวณชั่วโมงที่สามารถใช้งานได้สัมพันธ์กับจำนวนวันสุทธิ
merged["Available Hours"] = merged["Open Days"] * HOURS_PER_DAY

def calc_util_pct(row):
    avail = row["Available Hours"]
    dur = row["Total Duration (Hours)"]
    if avail > 0:
        val = (dur / avail) * 100
        return val / 100.0  
    return 0.0

merged["Utilization (%)"] = merged.apply(calc_util_pct, axis=1)

merged["File Date"] = pd.to_datetime(merged["File Date"], errors='coerce')
merged["From Year"] = merged["File Date"].dt.year
merged["From Month"] = merged["File Date"].dt.month
merged["From Day"] = merged["File Date"].dt.day
merged["To Year"] = merged["File Date"].dt.year
merged["To Month"] = merged["File Date"].dt.month
merged["To Day"] = merged["File Date"].dt.day

merged = merged.reindex(columns=FINAL_COLUMNS)

# สั่งเซฟ Debug Raw File 
debug_out = merged.copy()
debug_out["File Date"] = debug_out["File Date"].dt.strftime('%d/%m/%Y').fillna("Unknown")
debug_out["Utilization (%)"] = (debug_out["Utilization (%)"] * 100).apply(lambda x: f"{x:.1f}%")
debug_out.to_excel(DEBUG_RAW_FILE, index=False)

# ===================================================
# 3. เริ่มสร้างรายงานหลักและเปิดตารางรูปแบบ Excel แท้
# ===================================================
start_row = 3 
total_rows = len(merged)

def get_excel_col_name(col_index):
    result = ""
    while col_index >= 0:
        result = chr(int(col_index % 26) + 65) + result
        col_index = int(col_index / 26) - 1
    return result

col_map = {col: get_excel_col_name(i) for i, col in enumerate(FINAL_COLUMNS)}

with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
    merged.to_excel(writer, sheet_name="Summary by room", index=False, startrow=start_row)
    pd.DataFrame(debug_summary).to_excel(writer, sheet_name="Debug Log", index=False)

    workbook = writer.book
    worksheet = writer.sheets["Summary by room"]

    title_format = workbook.add_format({'font_name': 'Segoe UI', 'font_size': 18, 'bold': True, 'font_color': '#2C3E50'})
    date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'align': 'center'})
    pct_format = workbook.add_format({'num_format': '0.0%', 'align': 'right'})
    hour_num_format = workbook.add_format({'num_format': '0.0', 'align': 'right'})
    center_format = workbook.add_format({'align': 'center'})

    worksheet.write('A1', 'Summary by room', title_format)

    c_no = col_map["No"]
    c_date = col_map["File Date"]
    c_dur_h = col_map["Total Duration (Hours)"]
    c_days = col_map["Open Days"]
    c_avail = col_map["Available Hours"]
    c_util = col_map["Utilization (%)"]

    column_settings = [{'header': col} for col in FINAL_COLUMNS]
    worksheet.add_table(start_row, 0, start_row + total_rows, len(FINAL_COLUMNS) - 1, {
        'columns': column_settings,
        'style': 'Table Style Medium 9',
        'name': 'RoomSummaryTable'
    })

    worksheet.set_column(f'{c_dur_h}:{c_dur_h}', 14, hour_num_format)

    for i in range(total_rows):
        r_idx = start_row + 2 + i  
        
        actual_date = merged.loc[i, "File Date"]
        actual_days = int(merged.loc[i, "Open Days"])
        actual_avail = float(merged.loc[i, "Available Hours"])
        actual_util_val = float(merged.loc[i, "Utilization (%)"])
        
        worksheet.write(f'{c_no}{r_idx}', i + 1, center_format)
        
        for k in ["From Year", "From Month", "From Day", "To Year", "To Month", "To Day"]:
            val = merged.loc[i, k]
            if pd.notna(val):
                worksheet.write(f'{col_map[k]}{r_idx}', int(val), center_format)
            else:
                worksheet.write(f'{col_map[k]}{r_idx}', "", center_format)

        if pd.notna(actual_date):
            worksheet.write_datetime(f'{c_date}{r_idx}', actual_date, date_format)
        else:
            worksheet.write(f'{c_date}{r_idx}', "Unknown", center_format)
            
        worksheet.write(f'{c_days}{r_idx}', actual_days, center_format)
        worksheet.write(f'{c_avail}{r_idx}', actual_avail, hour_num_format)
        worksheet.write(f'{c_util}{r_idx}', actual_util_val, pct_format)

    worksheet.set_column('A:A', 6)    # No
    worksheet.set_column('B:B', 18)   # Building
    worksheet.set_column('C:H', 12)   # From Year -> To Day
    worksheet.set_column('I:N', 14)   # Room -> Total Duration
    worksheet.set_column('O:T', 15)   # Avg Duration -> Utilization (%)
    worksheet.set_column('U:W', 25)   # Source File, Source Sheet, File Date

print(f"🎉 รวมโค้ดสำเร็จ! ดักเงื่อนไขอาคาร THE TARA และห้องชั้นใต้ดิน (.B1) พร้อมแยกจำประเภทไฟล์เก่า-ใหม่เรียบร้อยครับพี่!")