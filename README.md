### Implements
- `bot.js`: เพิ่มหน้า pop-up กรอกรายละเอียด login, อาคาร, เลือกวันที่, จำนวนวัน
- ใช้ `merge.py`: filter building, วันเริ่ม-จบ ที่ดึงมาได้ + เปลี่ยนวันตามที่กรอกหน้า pop-up แต่ยัง set default บางชั้นตามที่แจ้งไว้ตอนแรก + ใส่ Total Duration (Hours), Available Hours, Utilization (%)
- `run_pipeline.py`: สั่งรัน bot แล้ว export ผลเข้า excel ครบถ้วน (เข้า `output\`)

---

#### Additional
- `template.xlsx`: template ของ excel ที่มี filter แบบ slicer ด้านบน
- `merge_xlwings.py`: รันเพื่อแสดงผลเป็น excel ที่มี slicer ขึ้นมาเลย (ดึงข้อมูลจาก `downloads\`)
- 
  ```
  pip install xlwings pandas openpyxl
  ```
  * แต่ใช้เปิดกับเครื่องที่มี excel ในตัวเท่านั้น
  * <img width="1828" height="305" alt="image" src="https://github.com/user-attachments/assets/741b4825-56c5-489f-8c51-a28bba1a99da" />
