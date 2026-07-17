### Implements
- หน้า pop-up กรอกรายละเอียด login, เลือกวันที่, จำนวนวัน, อาคาร
- รวมไฟล์ของห้อง ที่เลือกอาคารเดียวกัน ช่วงวันเดียวกัน
- ใส่ จำนวนชม., ชม.ที่ใช้คำนวณ, %ที่ได้จริง
- filter แต่ละ column
- bot.js: เพิ่มหน้า pop-up กรอกรายละเอียด login, อาคาร, เลือกวันที่, จำนวนวัน
- ใช้ merge.py: filter building, วันเริ่ม-จบ ที่ดึงมาได้ + เปลี่ยนวันตามที่กรอกหน้า pop-up แต่ยัง set default บางชั้นตามที่แจ้งไว้ตอนแรก + ใส่ Total Duration (Hours), Available Hours, Utilization (%)
- run_pipeline.py: สั่งรัน bot แล้ว export ผลเข้า excel ครบถ้วน (เข้า `output\`)

---

#### Additional
- template.xlsx: template ของ excel ที่มี filter แบบ slicer ด้านบน
- merge_xlwings.py: รันเพื่อแสดงผลเป็น excel ที่มี slicer ขึ้นมาเลย (ดึงข้อมูลจาก `downloads\`)
  ```
  pip install xlwings pandas openpyxl
  ```
  * แต่ใช้เปิดกับเครื่องที่มี excel ในตัวเท่านั้น