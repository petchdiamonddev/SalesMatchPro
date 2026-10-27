# SalesMatchPro

โปรแกรมเปรียบเทียบและจับคู่ข้อมูล Excel ระหว่างไฟล์รายการขายกับฐานข้อมูลหลัก โดยจะไฮไลท์แถวที่ตรงกัน (สีเหลือง) และลบแถวที่ไม่เกี่ยวข้องออก

## การจัดวางโฟลเดอร์

```text
SalesMatchPro/
├── SalesMatchPro.exe
├── input_sold/       <- ใส่ไฟล์รายการขาย (.xlsx)
├── input_all_data/   <- ใส่ไฟล์ฐานข้อมูลหลัก (.xlsx) 1 ไฟล์
└── output/           <- โฟลเดอร์เก็บไฟล์ผลลัพธ์
```

## วิธีใช้งาน

1. นำไฟล์ Excel ไปใส่ใน `input_sold/` และ `input_all_data/`
2. ดับเบิลคลิกไฟล์ `SalesMatchPro.exe` (หรือรัน `python process.py`)
3. รอประมวลผลเสร็จ ไฟล์ผลลัพธ์จะอยู่ในโฟลเดอร์ `output/`

### รันผ่าน Command Line (CLI)
```bash
python process.py
python process.py --column "รายการ" --no-pause
```

## การพัฒนาและ Build .exe

ติดตั้ง dependencies:
```bash
pip install -r requirements.txt
```

สร้างไฟล์ `.exe`:
```bash
build.bat
```
*(ไฟล์ผลลัพธ์จะอยู่ที่ `dist/SalesMatchPro.exe`)*
