# 📊 SalesMatchPro

> **SalesMatchPro** is an automated Excel data matching tool. It compares daily sold items against a master database, highlights matched rows in yellow, and automatically removes unmatched records. **No coding required!**

---

## 🚀 Quick Start (สำหรับผู้ใช้งาน)

เครื่องมือนี้ทำมาเพื่อให้ใช้งานได้ทันที (Plug & Play)

1. วางไฟล์ `.exe` และโฟลเดอร์ต่างๆ ไว้ด้วยกันตามนี้:
   ```text
   SalesMatchPro/
   ├── SalesMatchPro.exe
   ├── HOW_TO_USE.txt
   ├── input_sold/       <-- วางไฟล์ "รายการขาย" (ใส่ได้หลายไฟล์)
   ├── input_all_data/   <-- วางไฟล์ "ฐานข้อมูลหลัก" (ใส่ได้ 1 ไฟล์)
   └── output/           <-- ผลลัพธ์จะอยู่ที่นี่
   ```
2. ดับเบิลคลิก `SalesMatchPro.exe`
3. รอโปรแกรมประมวลผล แล้วเข้าไปดูไฟล์ใหม่ที่หน้าตาเปลี่ยนไปแล้วใน `output/`

---

## 💻 สำหรับนักพัฒนา (For Developers)

พัฒนาด้วย `Python` ร่วมกับ `pandas` และ `openpyxl` 

### ⚙️ ติดตั้งระบบ (Setup)
```bash
# 1. ติดตั้ง libraries ที่จำเป็น
pip install -r requirements.txt

# 2. รันโปรแกรมโดยตรง (ไม่ต้องผ่าน .exe)
python process.py
```

### 📦 การแพ็กเป็นไฟล์ `.exe` (Build Release)

หากต้องการ build ไฟล์ให้คนอื่นนำไปใช้งานต่อ โดยไม่ต้องลง Python:
```bash
# 1. ติดตั้ง PyInstaller
pip install pyinstaller

# 2. กดรันไฟล์ build.bat เพื่อสร้าง .exe อัตโนมัติ
# ไฟล์ผลลัพธ์จะถูกสร้างไปที่โฟลเดอร์ /dist
```

---

## 🎯 ฟีเจอร์หลัก (Key Features)

- **Auto-Detect Column:** ไม่จำเป็นต้องเรียงตารางเหมือนเดิมเป๊ะๆ โปรแกรมจะกวาดหาคอลัมน์ชื่อ "รายการ" เอง
- **Batch Processing:** ประมวลผล `input_sold/*.xlsx` ได้พร้อมกันเป็นสิบๆ ไฟล์ในการกดรันเพียงครั้งเดียว
- **Highlighting & Pruning:** ไฮไลท์ข้อมูลที่พบ (สีเหลือง) และลบขยะ/ข้อมูลที่ไม่เกี่ยวข้องออกทันที

---

## 📝 License

Distributed under the MIT License. See `LICENSE` for more information.
