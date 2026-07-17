import sys
import math
import re
import json
from pathlib import Path
import pandas as pd
import calendar
import xlwings as xw

DOWNLOAD_DIR = Path("downloads")
OUTPUT_DIR = Path("output")
OUTPUT_FILE = OUTPUT_DIR / "merged_report.xlsx"
DEBUG_RAW_FILE = OUTPUT_DIR / "debug_raw_concat.xlsx"
TEMPLATE_FILE = Path("template.xlsx")

OUTPUT_DIR.mkdir(exist_ok=True)

if not TEMPLATE_FILE.exists():
    print(f"❌ ไม่พบไฟล์แม่แบบ '{TEMPLATE_FILE.name}' กรุณาเตรียม Template ให้เรียบร้อยก่อนนะครับ")
    raise SystemExit

# ===================================================
# 1. ระบบรับค่าดักจับจาก Pipeline (Open Days & New Files)
# ===================================================
USER_NUM_DAYS = 19 
NEW_FILES_LIST = set()

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
        NEW_FILES_LIST = set(json.loads(sys.argv[-1]))
    except Exception:
        NEW_FILES_LIST = set()

HOURS_PER_DAY = 8.5

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
        if '_report_' in name: name = name.split('_report_')[0]
        else: name = name.rsplit('_', 1)[0]
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
    if start_idx > 0: df = df.iloc[start_idx:].copy().reset_index(drop=True)
    if df.shape[1] < 9: return None
    df = df.iloc[:, :9].copy()
    df.columns = EXPECTED_OUTPUT_COLUMNS
    header_mask = df.apply(lambda r: looks_like_header_row(r.tolist()), axis=1)
    df = df[~header_mask].copy()
    
    def clean_room_string(val):
        if pd.isna(val): return ""
        if isinstance(val, pd.Timestamp) or hasattr(val, 'strftime'):
            return val.strftime('%d') if val.day == 1 and val.month == 1 else val.strftime('%d-%m')
        val_str = str(val).strip()
        if "-" in val_str and "00:00:00" in val_str:
            val_str = val_str.split()[0]
            parts = val_str.split("-")
            if len(parts) == 3:
                if parts[1] == "01" and parts[2] == "01": return "01"
                return f"{parts[2]}-{parts[1]}"
        if val_str.endswith('.0'): return val_str[:-2]
        return val_str

    df["Room"] = df["Room"].apply(clean_room_string)
    invalid_keywords = ['undefined', 'undifine', 'total', 'nan', 'summary', '']
    df = df[df["Room"].notna() & (df["Room"] != "")]
    df = df[~df["Room"].str.lower().isin(invalid_keywords)].copy()
    df["Source File"] = source_file
    df["Source Sheet"] = source_sheet
    return df

# 💡 ฟังก์ชันคำนวณค่า Open Days พิเศษตามเงื่อนไขของพี่
def calculate_open_days(row, input_days_from_ui, new_files_set):
    building = str(row.get('Building', '')).strip()
    room = str(row.get('Room', '')).strip()
    source_file = str(row.get('Source File', '')).strip()
    
    # 📌 1. THE TARA ชั้น 22
    if building == "THE TARA" and room.startswith("22-"):
        return 22
        
    # 📌 2. ห้องกลุ่มชั้นใต้ดินเจาะจง (.B / .B1 / B)
    b1_rooms_mapping = {
        ".B1-01": 20, ".B1-02": 20, ".B1-03": 20, ".B1-04": 21,
        ".B1-05": 22, ".B1-06": 20, ".B1-07": 20, ".B1-08": 21
    }
    clean_room = room.replace(".", "").upper()
    for b1_key, days_val in b1_rooms_mapping.items():
        clean_key = b1_key.replace(".", "").upper()
        if clean_key == clean_room or b1_key in room.upper():
            return days_val

    # 📌 3. เช็กสถานะไฟล์ใหม่-เก่า
    if source_file in new_files_set:
        return input_days_from_ui
        
    existing_days = row.get('Open Days', None)
    if pd.notna(existing_days):
        try: return int(float(existing_days))
        except ValueError: pass
            
    return 19

# ===================================================
# ดึงข้อมูลดิบจาก downloads
# ===================================================
files = sorted([f for f in DOWNLOAD_DIR.glob("*.xlsx")])
if not files:
    print(f"❌ ไม่พบไฟล์ Excel ใด ๆ ใน {DOWNLOAD_DIR.absolute()}")
    raise SystemExit

all_data = []
debug_summary = []
detected_date_objects = [] 

for file in files:
    try:
        file_date = extract_date_from_filename(file.name)
        if file_date: detected_date_objects.append(file_date)
        xl = pd.ExcelFile(file)
        building_name = extract_building_name(file.name)
        for sheet_name in xl.sheet_names:
            raw_df = pd.read_excel(file, sheet_name=sheet_name, header=None, dtype=str)
            
            # ควานหาค่า Open Days ดั้งเดิมจากตารางเก่ากรณีไม่ใช่ไฟล์ใหม่
            existing_days_val = None
            if raw_df.shape[1] >= 18:
                for idx in range(min(15, len(raw_df))):
                    val = raw_df.iloc[idx, 17]
                    if pd.notna(val) and str(val).replace('.0', '').isdigit():
                        existing_days_val = int(str(val).replace('.0', ''))
                        break

            normalized_df = normalize_sheet(raw_df, file.name, sheet_name)
            if normalized_df is not None and not normalized_df.empty:
                normalized_df["Building"] = building_name
                normalized_df["File Date"] = file_date  
                normalized_df["Open Days"] = existing_days_val
                all_data.append(normalized_df)
                debug_summary.append({"file": file.name, "sheet": sheet_name, "rows": len(normalized_df), "status": "OK"})
    except Exception as e:
        debug_summary.append({"file": file.name, "sheet": str(e), "rows": 0, "status": "Error"})

if not all_data:
    print("❌ 沒有ข้อมูลให้ประมวลผล")
    raise SystemExit

merged = pd.concat(all_data, ignore_index=True)
invalid_global = ['undefined', 'undifine', 'total', 'summary', '']
merged = merged[~merged["Room"].astype(str).str.strip().str.lower().isin(invalid_global)].copy()

merged["_sort"] = merged["Room"].apply(parse_room_key)
merged = merged.sort_values(by=["Building", "_sort"], kind="stable").drop(columns="_sort").reset_index(drop=True)
merged["Total Duration (Hours)"] = merged["Total Duration"].apply(parse_duration_to_hours)

merged["File Date"] = pd.to_datetime(merged["File Date"], errors='coerce')
merged["From Year"] = merged["File Date"].dt.year
merged["From Month"] = merged["File Date"].dt.month
merged["From Day"] = merged["File Date"].dt.day
merged["To Year"] = merged["File Date"].dt.year
merged["To Month"] = merged["File Date"].dt.month
merged["To Day"] = merged["File Date"].dt.day

merged["No"] = range(1, len(merged) + 1)

# 🔥 สั่งคำนวณคอลัมน์ Open Days ใน Dataframe ให้เสร็จสรรพตามเงื่อนไขพิเศษรายห้อง
merged["Open Days"] = merged.apply(lambda r: calculate_open_days(r, USER_NUM_DAYS, NEW_FILES_LIST), axis=1)

merged["Available Hours"] = 0.0
merged["Utilization (%)"] = 0.0
merged = merged.reindex(columns=FINAL_COLUMNS)

debug_out = merged.copy()
debug_out["File Date"] = debug_out["File Date"].dt.strftime('%d/%m/%Y').fillna("Unknown")
debug_out.to_excel(DEBUG_RAW_FILE, index=False)

def get_excel_col_name(col_index):
    result = ""
    while col_index >= 0:
        result = chr(int(col_index % 26) + 65) + result
        col_index = int(col_index / 26) - 1
    return result

col_map = {col: get_excel_col_name(i) for i, col in enumerate(FINAL_COLUMNS)}

# ===================================================
# 🚀 เริ่มเขียนข้อมูลผ่าน xlwings 
# ===================================================
print("🚀 กำลังสตาร์ทโปรแกรม Microsoft Excel ผ่าน xlwings และรักษาเลขห้อง...")

app = xw.App(visible=False)
try:
    wb = app.books.open(TEMPLATE_FILE)
    sheet_name_map = {s.name.strip().lower(): s.name for s in wb.sheets}

    def write_data_xlwings(sheet_obj, df_data, title_name):
        sheet_obj.range('A1').value = title_name
        sheet_obj.range('D1').value = "Open Days ต้นแบบ:"
        sheet_obj.range('E1').value = USER_NUM_DAYS
        
        start_row = 12
        t_rows = len(df_data)
        
        table_found = None
        native_ws = sheet_obj.api  
        try:
            if hasattr(native_ws, 'ListObjects') and native_ws.ListObjects.Count > 0:
                table_found = native_ws.ListObjects(1)
        except Exception:
            pass

        last_row = sheet_obj.range('A' + str(sheet_obj.cells.last_cell.row)).end('up').row
        if last_row > start_row:
            sheet_obj.range(f"A{start_row+1}:{get_excel_col_name(len(FINAL_COLUMNS)-1)}{last_row}").clear_contents()

        for i in range(t_rows):
            r_idx = start_row + 1 + i
            actual_date = df_data.loc[i, "File Date"]
            actual_dur_h = float(df_data.loc[i, "Total Duration (Hours)"])
            final_calculated_days = int(df_data.loc[i, "Open Days"]) # ดึงค่าที่ผ่านการคำนวณเงื่อนไขพิเศษมาใช้
            
            sheet_obj.range(f'{col_map["No"]}{r_idx}').value = i + 1
            sheet_obj.range(f'{col_map["Building"]}{r_idx}').value = df_data.loc[i, "Building"]
            
            for k in ["From Year", "From Month", "From Day", "To Year", "To Month", "To Day"]:
                val = df_data.loc[i, k]
                sheet_obj.range(f'{col_map[k]}{r_idx}').value = "" if pd.isna(val) or math.isnan(val) else int(val)

            room_cell = sheet_obj.range(f'{col_map["Room"]}{r_idx}')
            room_cell.number_format = '@'
            room_cell.value = str(df_data.loc[i, "Room"])

            sheet_obj.range(f'{col_map["Number of Bookings"]}{r_idx}').value = df_data.loc[i, "Number of Bookings"]
            sheet_obj.range(f'{col_map["Number of Meet Now"]}{r_idx}').value = df_data.loc[i, "Number of Meet Now"]
            sheet_obj.range(f'{col_map["Number of Advanced Bookings"]}{r_idx}').value = df_data.loc[i, "Number of Advanced Bookings"]
            sheet_obj.range(f'{col_map["Total Duration"]}{r_idx}').value = df_data.loc[i, "Total Duration"]
            sheet_obj.range(f'{col_map["Total Duration (Hours)"]}{r_idx}').value = actual_dur_h
            sheet_obj.range(f'{col_map["Avg Duration"]}{r_idx}').value = df_data.loc[i, "Avg Duration"]
            sheet_obj.range(f'{col_map["Min Duration"]}{r_idx}').value = df_data.loc[i, "Min Duration"]
            sheet_obj.range(f'{col_map["Max Duration"]}{r_idx}').value = df_data.loc[i, "Max Duration"]
            
            # 💡 ไฮไลต์แก้วิธีหยอดตรงนี้: 
            # ถ้าเป็นห้องทั่วไปให้ผูกสูตรไปหาช่อง E1 ตามปกติ แต่ถ้าเป็นกรณีเงื่อนไขพิเศษ ให้หยอดตัวเลขดิบ (20,21,22) ทับสูตรลงไปดื้อๆ เลย
            building_name = str(df_data.loc[i, "Building"]).strip()
            room_name = str(df_data.loc[i, "Room"]).strip()
            
            is_special_tara = (building_name == "THE TARA" and room_name.startswith("22-"))
            is_special_b1 = any(clean_k in room_name.replace(".", "").upper() for clean_k in ["B1-01", "B1-02", "B1-03", "B1-04", "B1-05", "B1-06", "B1-07", "B1-08"])
            
            if is_special_tara or is_special_b1:
                sheet_obj.range(f'{col_map["Open Days"]}{r_idx}').value = final_calculated_days
            else:
                sheet_obj.range(f'{col_map["Open Days"]}{r_idx}').value = "=E$1"
                
            sheet_obj.range(f'{col_map["Available Hours"]}{r_idx}').value = f'={col_map["Open Days"]}{r_idx} * {HOURS_PER_DAY}'
            
            util_cell = sheet_obj.range(f'{col_map["Utilization (%)"]}{r_idx}')
            util_cell.value = f'=IF({col_map["Available Hours"]}{r_idx}>0, {col_map["Total Duration (Hours)"]}{r_idx}/{col_map["Available Hours"]}{r_idx}, 0)'
            util_cell.number_format = '0.0%'
            
            sheet_obj.range(f'{col_map["Source File"]}{r_idx}').value = df_data.loc[i, "Source File"]
            sheet_obj.range(f'{col_map["Source Sheet"]}{r_idx}').value = df_data.loc[i, "Source Sheet"]
            
            if pd.notna(actual_date):
                sheet_obj.range(f'{col_map["File Date"]}{r_idx}').value = actual_date.strftime('%d/%m/%Y')
            else:
                sheet_obj.range(f'{col_map["File Date"]}{r_idx}').value = "Unknown"

        if table_found:
            new_ref = f"A12:{get_excel_col_name(len(FINAL_COLUMNS)-1)}{max(start_row+1, start_row + t_rows)}"
            try:
                table_found.Resize(native_ws.Range(new_ref))
            except Exception as e:
                print(f"⚠️ ไม่สามารถอัปเดตขนาดตารางโดยอัตโนมัติ: {e}")

    # 1. เขียนชีทรวม Summary by room
    target_summary_key = "summary by room"
    if target_summary_key in sheet_name_map:
        write_data_xlwings(wb.sheets[sheet_name_map[target_summary_key]], merged, "Summary by room (All Buildings)")
    else:
        write_data_xlwings(wb.sheets[0], merged, "Summary by room (All Buildings)")

    # 2. เขียนชีทรายอาคาร
    unique_buildings = merged["Building"].unique()
    for b_name in unique_buildings:
        b_df = merged[merged["Building"] == b_name].reset_index(drop=True)
        sheet_title = str(b_name)[:30].strip().lower()
        
        matched_sheet = None
        for s in wb.sheets:
            if s.name.strip().lower() == sheet_title:
                matched_sheet = s
                break
                
        if matched_sheet:
            print(f"📦 เขียนชีทตึก: {matched_sheet.name}")
            write_data_xlwings(matched_sheet, b_df, f"Summary by room - {b_name}")
        else:
            print(f"⚠️ ไม่พบชีทชื่อตึก '{b_name}' ในเทมเพลต")

    wb.save(OUTPUT_FILE)
    wb.close()
    print("🎉 ดำเนินการเสร็จสิ้นสมบูรณ์! รองรับระบบแยกประเภทไฟล์เก่า-ใหม่และดักเงื่อนไขพิเศษของตึก THE TARA และห้องชั้นใต้ดิน (.B1) บนระบบ xlwings เรียบร้อยครับพี่!")

finally:
    app.quit()