import subprocess
import sys
import re
import os
import json
from pathlib import Path

# ⚠️ ตั้งค่าโฟลเดอร์ดาวน์โหลดให้ตรงกับใน bot.js (ปกติคือ './downloads')
DOWNLOAD_DIR = Path("./downloads")
OUTPUT_FILE = Path("output/merged_report.xlsx")

def get_existing_files():
    """ฟังก์ชันเช็กรายชื่อไฟล์ Excel ที่มีอยู่แล้วในโฟลเดอร์ downloads"""
    if not DOWNLOAD_DIR.exists():
        return set()
    return {f.name for f in DOWNLOAD_DIR.glob("*.xlsx")}

def clear_excel_file_if_locked(file_path):
    """ฟังก์ชันล้างไฟล์ผลลัพธ์หลัก ป้องกัน Permission Error"""
    if file_path.exists():
        try:
            os.remove(file_path)
            print(f"🗑️  [Pipeline] ลบไฟล์รายงานผลลัพธ์เดิม ({file_path.name}) เพื่อเตรียมเขียนทับ...")
        except PermissionError:
            if os.name == 'nt':
                try:
                    subprocess.run(["taskkill", "/f", "/im", "excel.exe"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    os.remove(file_path)
                    print(f"✅ [Pipeline] ปลดล็อกและลบไฟล์ {file_path.name} สำเร็จ!")
                except Exception:
                    print(f"⚠️  [Pipeline] ไม่สามารถลบไฟล์ผลลัพธ์ได้เนื่องจากถูกล็อกอย่างรุนแรง")

def main():
    # 1. บันทึกรายชื่อไฟล์ที่มีอยู่แล้ว "ก่อนรันบอท"
    files_before = get_existing_files()
    
    print("🤖 [Pipeline] ขั้นตอนที่ 1: กำลังเริ่มรันบอท (กรอกข้อมูลที่หน้าต่างสีฟ้าได้เลย)...")
    
    # รันบอท bot.js
    process = subprocess.Popen(
        ["node", "bot.js"], 
        stdout=subprocess.PIPE, 
        stderr=subprocess.STDOUT, 
        text=True, 
        encoding='utf-8'
    )
    
    detected_days = None
    
    # อ่าน Log ของบอทแบบ Real-time
    while True:
        output = process.stdout.readline()
        if output == '' and process.poll() is not None:
            break
        if output:
            line = output.strip()
            print(line)
            
            # แกะตัวเลขวันจากหน้าจอที่กรอกใหม่
            if "จำนวนวันที่บันทึก" in line:
                match = re.search(r"(\d+)\s*วัน", line)
                if match:
                    detected_days = int(match.group(1))
                    print(f"🎯 [Pipeline] ตรวจพบจำนวนวันที่กรอกใหม่: {detected_days} วัน")

    process.wait()

    if process.returncode != 0:
        print("❌ [Pipeline] เกิดข้อผิดพลาดในการทำงานของบอท ระบบหยุดทำงาน")
        return

    # 2. ตรวจสอบไฟล์ที่มี "หลังจากบอทรันเสร็จ" เพื่อหาว่าไฟล์ไหนเพิ่งโหลดมาใหม่
    files_after = get_existing_files()
    new_files = files_after - files_before
    
    # ค่าเริ่มต้นถ้าบอทแกะเลขไม่ได้ หรือสำหรับไฟล์ที่จำประวัติไม่ได้เลย ให้เป็น 19 วันตามที่พี่สั่งครับ
    input_days = detected_days if detected_days is not None else 19
    
    print(f"\n📊 [Pipeline] ขั้นตอนที่ 2: จำแนกสถานะไฟล์ดาวน์โหลดเรียบร้อย")
    if new_files:
        print(f"✨ พบไฟล์ที่ดาวน์โหลดมาใหม่ในรอบนี้จำนวน {len(new_files)} ไฟล์ (จะใช้ค่า {input_days} วัน):")
        for f in new_files:
            print(f"   - {f}")
    else:
        print("ℹ️ ไม่พบไฟล์ดาวน์โหลดใหม่ในรอบนี้ (ระบบจะอ้างอิงข้อมูลเดิมของไฟล์เก่า)")

    # เคลียร์ไฟล์ล็อกตัวหลัก
    clear_excel_file_if_locked(OUTPUT_FILE)
    
    print(f"\n🚀 กำลังส่งไม้ต่อให้สคริปต์ Excel (merge.py)...")
    
    # 3. สั่งรัน merge.py 
    # โดยเราจะส่งค่าส่งท้ายไป 2 ค่า: ค่าวันใหม่ (input_days) และรายชื่อไฟล์ใหม่ (ส่งเป็น JSON string)
    # เพื่อให้ฝั่ง Excel แยกแยะได้ว่าไฟล์ไหนควรใช้ค่าเก่า ไฟล์ไหนควรใช้ค่าใหม่
    new_files_json = json.dumps(list(new_files))
    
    subprocess.run([
        sys.executable, 
        "merge.py", 
        str(input_days),        # sys.argv[-2] (จำนวนวันใหม่)
        new_files_json          # sys.argv[-1] (รายชื่อไฟล์ใหม่)
    ])
    
    print("\n🎉 ทุกขั้นตอนเสร็จสิ้นสมบูรณ์แบบ แยกประเภทไฟล์เก่า-ใหม่เรียบร้อยครับพี่!")

if __name__ == "__main__":
    main()