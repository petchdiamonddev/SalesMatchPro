"""
โปรแกรมประมวลผลไฟล์ Excel
(Excel Processing Program)

วิธีใช้งาน (How to use):
1. ใส่ไฟล์รายการที่ขายใน folder 'input_sold/'
2. ใส่ไฟล์ข้อมูลการขายทั้งหมดใน folder 'input_all_data/'
3. รันโปรแกรม: python process.py
4. ผลลัพธ์จะอยู่ใน folder 'output/'
"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import os
import sys
from pathlib import Path

class ExcelProcessor:
    def __init__(self):
        """เริ่มต้นโปรแกรม (Initialize program)"""
        # รองรับทั้ง .py และ .exe (Support both .py and .exe)
        if getattr(sys, 'frozen', False):
            self.base_dir = Path(sys.executable).parent
        else:
            self.base_dir = Path(__file__).parent
        self.input_sold_dir = self.base_dir / "input_sold"
        self.input_all_data_dir = self.base_dir / "input_all_data"
        self.output_dir = self.base_dir / "output"
        
        # สีเหลืองสำหรับไฮไลท์ (Yellow color for highlighting)
        self.yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    
    def create_folders(self):
        """สร้าง folder ถ้ายังไม่มี (Create folders if not exist)"""
        self.input_sold_dir.mkdir(exist_ok=True)
        self.input_all_data_dir.mkdir(exist_ok=True)
        self.output_dir.mkdir(exist_ok=True)
        print("✅ สร้าง folders เรียบร้อย (Folders created)")
    
    def find_column_by_header(self, df, header_name):
        """
        หาคอลัมน์จากชื่อหัวตาราง (Find column by header name)
        
        Parameters:
            df: DataFrame
            header_name: ชื่อหัวตารางที่ต้องการหา (Header name to find)
        
        Returns:
            column_index: ตำแหน่งคอลัมน์ (Column index) หรือ None ถ้าไม่เจอ
        """
        for idx, col in enumerate(df.columns):
            if header_name.lower() in str(col).lower():
                print(f"   ✅ พบคอลัมน์ '{header_name}' ที่ตำแหน่ง: {idx} (ชื่อจริง: {col})")
                return idx
        return None
    
    def get_all_data_file(self):
        """
        หาไฟล์ข้อมูลการขายทั้งหมด (Find the all data file)
        
        Returns:
            Path object หรือ None ถ้าไม่เจอ
        """
        xlsx_files = list(self.input_all_data_dir.glob("*.xlsx"))
        
        if len(xlsx_files) == 0:
            return None
        elif len(xlsx_files) > 1:
            print(f"⚠️  พบหลายไฟล์ใน input_all_data/ จะใช้ไฟล์แรก: {xlsx_files[0].name}")
        
        return xlsx_files[0]
    
    def process_single_file(self, sold_file_path, all_data_file_path):
        """
        ประมวลผลไฟล์เดี่ยว (Process single file)
        
        Parameters:
            sold_file_path: ไฟล์รายการที่ขาย (Sold items file)
            all_data_file_path: ไฟล์ข้อมูลทั้งหมด (All data file)
        
        Returns:
            (success, message, output_filename)
        """
        try:
            print(f"\n{'='*60}")
            print(f"📄 กำลังประมวลผล: {sold_file_path.name}")
            print(f"{'='*60}")
            
            # --- Part 1: อ่านไฟล์รายการที่ขาย ---
            print("📖 [1/5] กำลังอ่านไฟล์รายการที่ขาย...")
            df_sold = pd.read_excel(sold_file_path, header=1)
            
            # หาคอลัมน์ "รายการ" อัตโนมัติ (Auto-detect "รายการ" column)
            sold_col_idx = self.find_column_by_header(df_sold, "รายการ")
            if sold_col_idx is None:
                return False, "❌ ไม่พบคอลัมน์ 'รายการ' ในไฟล์รายการที่ขาย", None
            
            # ดึงรายการทั้งหมดออกมา (Extract all items)
            ids_to_find = set(df_sold.iloc[:, sold_col_idx].dropna())
            print(f"   ✅ พบ {len(ids_to_find)} รายการที่ต้องค้นหา")
            
            # --- Part 2: โหลดไฟล์ข้อมูลทั้งหมด ---
            print("📖 [2/5] กำลังโหลดไฟล์ข้อมูลทั้งหมด...")
            df_all = pd.read_excel(all_data_file_path, header=1)
            
            # หาคอลัมน์ "รายการ" ในไฟล์ข้อมูลทั้งหมด
            all_col_idx = self.find_column_by_header(df_all, "รายการ")
            if all_col_idx is None:
                return False, "❌ ไม่พบคอลัมน์ 'รายการ' ในไฟล์ข้อมูลทั้งหมด", None
            
            # โหลด workbook ด้วย openpyxl เพื่อแก้ไข
            wb = load_workbook(all_data_file_path)
            ws = wb.active
            print("   ✅ โหลดไฟล์สำเร็จ")
            
            # --- Part 3: ไฮไลท์และเก็บแถวที่จะลบ ---
            print("🔍 [3/5] กำลังตรวจสอบและไฮไลท์แถว...")
            rows_to_delete = []
            highlighted_count = 0
            
            # วนลูปตรวจสอบแต่ละแถว (min_row=3 เพราะมี header ที่แถว 2)
            for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
                # ดึงค่าจากคอลัมน์ที่เจอ
                cell_id = row[all_col_idx].value
                
                if cell_id in ids_to_find:
                    # ถ้าพบรายการตรงกัน ให้ไฮไลท์ทั้งแถว (Highlight matched rows)
                    for cell in row:
                        cell.fill = self.yellow_fill
                    highlighted_count += 1
                else:
                    # ถ้าไม่ตรง ให้เก็บไว้เพื่อลบ (Keep for deletion)
                    rows_to_delete.append(row[0].row)
            
            print(f"   ✅ ไฮไลท์ {highlighted_count} แถว")
            print(f"   📝 พบ {len(rows_to_delete)} แถวที่จะลบ")
            
            # --- Part 4: ลบแถวที่ไม่ตรงกัน ---
            if rows_to_delete:
                print("🗑️  [4/5] กำลังลบแถวที่ไม่ตรงกัน...")
                # ลบจากล่างขึ้นบน เพื่อไม่ให้ index เปลี่ยน (Delete from bottom to top)
                for row_index in sorted(rows_to_delete, reverse=True):
                    ws.delete_rows(row_index)
                print(f"   ✅ ลบ {len(rows_to_delete)} แถวเรียบร้อย")
            else:
                print("   ℹ️  ไม่มีแถวที่ต้องลบ")
            
            # --- Part 5: บันทึกไฟล์ ---
            print("💾 [5/5] กำลังบันทึกไฟล์...")
            
            # สร้างชื่อไฟล์ output (Create output filename)
            original_name = sold_file_path.stem  # ชื่อไฟล์ไม่มี .xlsx
            output_filename = self.output_dir / f"{original_name}_processed.xlsx"
            
            wb.save(output_filename)
            print(f"   ✅ บันทึกเป็น: {output_filename.name}")
            
            return True, f"✅ ประมวลผลสำเร็จ! ({highlighted_count} รายการตรงกัน, ลบ {len(rows_to_delete)} แถว)", output_filename.name
            
        except Exception as e:
            return False, f"❌ เกิดข้อผิดพลาด: {str(e)}", None
    
    def run(self):
        """เริ่มประมวลผลทั้งหมด (Start processing all files)"""
        print("\n" + "="*60)
        print("🚀 เริ่มโปรแกรมประมวลผล Excel")
        print("   (Excel Processing Program Started)")
        print("="*60)
        
        # สร้าง folders
        self.create_folders()
        
        # หาไฟล์ข้อมูลทั้งหมด (Find all data file)
        all_data_file = self.get_all_data_file()
        if all_data_file is None:
            print("\n❌ ข้อผิดพลาด: ไม่พบไฟล์ใน folder 'input_all_data/'")
            print("   กรุณาใส่ไฟล์ 'ข้อมูลการขายทั้งหมด.xlsx' ใน folder นี้")
            return
        
        print(f"\n📁 ใช้ไฟล์ข้อมูลทั้งหมด: {all_data_file.name}")
        
        # หาไฟล์รายการที่ขายทั้งหมด (Find all sold item files)
        sold_files = list(self.input_sold_dir.glob("*.xlsx"))
        
        if len(sold_files) == 0:
            print("\n❌ ข้อผิดพลาด: ไม่พบไฟล์ใน folder 'input_sold/'")
            print("   กรุณาใส่ไฟล์รายการที่ขาย (.xlsx) ใน folder นี้")
            return
        
        print(f"\n📋 พบ {len(sold_files)} ไฟล์ที่จะประมวลผล:")
        for f in sold_files:
            print(f"   • {f.name}")
        
        # ประมวลผลแต่ละไฟล์ (Process each file)
        results = []
        for sold_file in sold_files:
            success, message, output_name = self.process_single_file(sold_file, all_data_file)
            results.append({
                'file': sold_file.name,
                'success': success,
                'message': message,
                'output': output_name
            })
        
        # สรุปผลลัพธ์ (Summary)
        print("\n" + "="*60)
        print("📊 สรุปผลการประมวลผล (Summary)")
        print("="*60)
        
        success_count = sum(1 for r in results if r['success'])
        fail_count = len(results) - success_count
        
        for result in results:
            status = "✅" if result['success'] else "❌"
            print(f"{status} {result['file']}")
            print(f"   {result['message']}")
            if result['output']:
                print(f"   📄 ไฟล์ผลลัพธ์: {result['output']}")
        
        print(f"\n✅ สำเร็จ: {success_count} ไฟล์")
        if fail_count > 0:
            print(f"❌ ล้มเหลว: {fail_count} ไฟล์")
        
        print("\n🎉 เสร็จสิ้นการประมวลผล!")
        print(f"📁 ผลลัพธ์อยู่ใน folder: {self.output_dir}")
        print("="*60 + "\n")

def main():
    """ฟังก์ชันหลัก (Main function)"""
    try:
        processor = ExcelProcessor()
        processor.run()
    except KeyboardInterrupt:
        print("\n\n⚠️  ยกเลิกการทำงานโดยผู้ใช้")
    except Exception as e:
        print(f"\n❌ เกิดข้อผิดพลาดร้ายแรง: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # รอกด Enter ก่อนปิด เพื่อให้ผู้ใช้เห็นผลลัพธ์ (Pause so user can see results)
        input("\nกด Enter เพื่อปิดโปรแกรม (Press Enter to close)...")

if __name__ == "__main__":
    main()
