import os
import tempfile
from pathlib import Path
import shutil
import subprocess
import re
import platform
from PyPDF2 import PdfMerger
from datetime import datetime

def find_and_compress_pdfs(input_directory, output_directory, dpi=150, image_quality=85):
    """
    ค้นหาไฟล์ PDF ทั้งหมดในไดเรกทอรีที่กำหนด และลดขนาดไฟล์โดยลดความละเอียดของรูปภาพภายใน
    
    Parameters:
    input_directory (str): ไดเรกทอรีที่ต้องการค้นหาไฟล์ PDF
    output_directory (str): ไดเรกทอรีที่ต้องการบันทึกไฟล์ PDF ที่ถูกบีบอัดแล้ว
    dpi (int): ความละเอียดของรูปภาพที่ต้องการ (dot per inch)
    image_quality (int): คุณภาพของรูปภาพ (0-100)
    
    Returns:
    dict: ข้อมูลสรุปเกี่ยวกับไฟล์ที่ถูกบีบอัด
    """
    # ตรวจสอบว่าเป็นระบบปฏิบัติการอะไรและเลือกคำสั่ง Ghostscript ที่เหมาะสม
    system = platform.system()
    if system == "Windows":
        # ลองทั้ง gswin64c และ gswin32c สำหรับ Windows
        gs_commands = ["gswin64c", "gswin32c", "gs"]
    else:
        # สำหรับ Linux และ macOS
        gs_commands = ["gs"]
    
    # ตรวจสอบว่ามี Ghostscript ติดตั้งแล้ว
    has_ghostscript = False
    gs_command = None
    
    for cmd in gs_commands:
        try:
            # ทดสอบเรียกใช้ Ghostscript
            result = subprocess.run([cmd, '--version'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if result.returncode == 0:
                has_ghostscript = True
                gs_command = cmd
                print(f"พบ Ghostscript ({cmd}) เวอร์ชัน: {result.stdout.decode().strip()}")
                break
        except (subprocess.SubprocessError, FileNotFoundError):
            continue
    
    if not has_ghostscript:
        print("ไม่พบ Ghostscript ที่ติดตั้ง จะใช้วิธีอื่นแทน")
    
    # สร้างไดเรกทอรีเอาท์พุทหากยังไม่มี
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    
    # ข้อมูลสรุป
    summary = {
        'total_files': 0,
        'compressed_files': 0,
        'failed_files': 0,
        'original_size': 0,
        'compressed_size': 0
    }
    
    # ค้นหาไฟล์ PDF ทั้งหมด
    pdf_files = []
    for root, dirs, files in os.walk(input_directory):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    
    summary['total_files'] = len(pdf_files)
    
    # ประมวลผลแต่ละไฟล์
    for pdf_file in pdf_files:
        try:
            # สร้างเส้นทางไฟล์เอาท์พุทโดยการรักษาโครงสร้างโฟลเดอร์เดิม
            rel_path = os.path.relpath(pdf_file, input_directory)
            output_path = os.path.join(output_directory, rel_path)
            
            # สร้างไดเรกทอรีย่อยถ้าจำเป็น
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            original_size = os.path.getsize(pdf_file)
            summary['original_size'] += original_size
            
            # วิธีที่ 1: ใช้ Ghostscript (ถ้ามี)
            if has_ghostscript:
                # คำสั่ง Ghostscript สำหรับลดขนาดรูปภาพใน PDF
                gs_params = [
                    gs_command, 
                    '-sDEVICE=pdfwrite', 
                    '-dCompatibilityLevel=1.4',
                    '-dPDFSETTINGS=/default',
                    f'-dDownsampleColorImages=true',
                    f'-dColorImageResolution={dpi}',
                    f'-dColorImageDownsampleThreshold=1.0',
                    f'-dDownsampleGrayImages=true',
                    f'-dGrayImageResolution={dpi}',
                    f'-dGrayImageDownsampleThreshold=1.0',
                    f'-dDownsampleMonoImages=true',
                    f'-dMonoImageResolution={dpi}',
                    f'-dMonoImageDownsampleThreshold=1.0',
                    f'-dJPEGQ={image_quality}',
                    '-dNOPAUSE',
                    '-dQUIET',
                    '-dBATCH',
                    f'-sOutputFile={output_path}',
                    pdf_file
                ]
                
                print(f"กำลังบีบอัดไฟล์: {pdf_file} ใช้ Ghostscript...")
                subprocess.run(gs_params, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
            
            # วิธีที่ 2: ใช้ pikepdf (หากติดตั้งแล้ว)
            else:
                try:
                    import pikepdf
                    from PIL import Image
                    
                    print(f"กำลังบีบอัดไฟล์: {pdf_file} ใช้ pikepdf...")
                    
                    # ใช้ pikepdf บีบอัดรูปภาพใน PDF
                    with pikepdf.open(pdf_file) as pdf:
                        # สำหรับแต่ละหน้าใน PDF
                        for page in pdf.pages:
                            # สำหรับแต่ละออบเจกต์ในหน้า
                            for name, obj in page.resources.get("/XObject", {}).items():
                                # ถ้าออบเจกต์เป็นรูปภาพ
                                if isinstance(obj, pikepdf.PdfImage):
                                    try:
                                        # แปลงเป็นรูปภาพ PIL
                                        img = obj.as_pil_image()
                                        
                                        # คำนวณขนาดใหม่ตาม DPI
                                        orig_dpi = obj.dpi or (300, 300)  # ถ้าไม่มี dpi ให้ใช้ 300 เป็นค่าเริ่มต้น
                                        width, height = img.size
                                        
                                        # ถ้า dpi ปัจจุบันมากกว่าที่ต้องการ ให้ปรับขนาดลง
                                        scale_factor = min(dpi / orig_dpi[0], dpi / orig_dpi[1])
                                        if scale_factor < 1:  # ลดขนาดเฉพาะเมื่อต้องการความละเอียดที่ต่ำกว่า
                                            new_width = int(width * scale_factor)
                                            new_height = int(height * scale_factor)
                                            img_resized = img.resize((new_width, new_height), Image.LANCZOS)
                                            
                                            # บันทึกรูปภาพที่ปรับขนาดแล้วเป็นไฟล์ชั่วคราว
                                            with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as tmp:
                                                tmp_path = tmp.name
                                                img_resized.save(tmp_path, format='JPEG', quality=image_quality, optimize=True)
                                            
                                            # แทนที่รูปภาพเดิมด้วยรูปภาพที่บีบอัดแล้ว
                                            new_image = pikepdf.PdfImage.open(tmp_path)
                                            page.resources["/XObject"][name] = pdf.add_image(new_image)
                                            
                                            # ลบไฟล์ชั่วคราว
                                            os.unlink(tmp_path)
                                    except Exception as e:
                                        print(f"  ไม่สามารถประมวลผลรูปภาพ: {e}")
                        
                        # บันทึก PDF ที่บีบอัดแล้ว
                        pdf.save(output_path,
                                compress_streams=True,
                                object_stream_mode=pikepdf.ObjectStreamMode.generate,
                                normalize_content=True)
                
                except ImportError:
                    print("ไม่พบ pikepdf และ Pillow ที่ติดตั้ง คัดลอกไฟล์โดยไม่บีบอัด")
                    shutil.copy2(pdf_file, output_path)
            
            compressed_size = os.path.getsize(output_path)
            summary['compressed_size'] += compressed_size
            summary['compressed_files'] += 1
            
            print(f"  ขนาดเดิม: {original_size/1024:.2f} KB")
            print(f"  ขนาดหลังบีบอัด: {compressed_size/1024:.2f} KB")
            print(f"  ลดลง: {(original_size - compressed_size) / original_size * 100:.2f}%")
            
        except Exception as e:
            print(f"เกิดข้อผิดพลาดในการบีบอัดไฟล์ {pdf_file}: {str(e)}")
            summary['failed_files'] += 1
            
            # ถ้าการบีบอัดล้มเหลว ให้ทำสำเนาไฟล์ต้นฉบับ (ทางเลือก)
            try:
                shutil.copy2(pdf_file, output_path)
                print(f"  คัดลอกไฟล์ต้นฉบับแทน")
            except:
                print(f"  ไม่สามารถคัดลอกไฟล์ต้นฉบับได้")
    
    # แสดงสรุป
    if summary['total_files'] > 0:
        print("\nสรุปการบีบอัด:")
        print(f"จำนวนไฟล์ทั้งหมด: {summary['total_files']}")
        print(f"จำนวนไฟล์ที่บีบอัดสำเร็จ: {summary['compressed_files']}")
        print(f"จำนวนไฟล์ที่บีบอัดล้มเหลว: {summary['failed_files']}")
        if summary['original_size'] > 0:
            savings = (1 - summary['compressed_size'] / summary['original_size']) * 100
            print(f"ขนาดเดิมรวม: {summary['original_size']/1024/1024:.2f} MB")
            print(f"ขนาดหลังบีบอัดรวม: {summary['compressed_size']/1024/1024:.2f} MB")
            print(f"ลดลงโดยเฉลี่ย: {savings:.2f}%")
    
    return summary

def merge_pdfs(input_directory, output_directory):
    pdf_merger = PdfMerger()
    
    # หามาไฟล์ PDF ทั้งหมดใน directory และเรียงตามชื่อ
    pdf_files = [file_name for file_name in os.listdir(input_directory) if file_name.endswith(".pdf")]
    pdf_files.sort()  # เรียงไฟล์ตามชื่อ
    
    for file_name in pdf_files:
        file_path = os.path.join(input_directory, file_name)
        pdf_merger.append(file_path)
    
    # สร้างชื่อไฟล์ output พร้อมวันที่
    date_str = datetime.now().strftime("%Y-%m-%d")  # วันที่ในรูปแบบปี-เดือน-วัน
    output_pdf = os.path.join(output_directory, f"Ex_{date_str}_merged.pdf")
    
    # สร้างไฟล์ PDF ที่รวม
    with open(output_pdf, "wb") as output_file:
        pdf_merger.write(output_file)
    print(f"ไฟล์ PDF ถูกรวมเรียบร้อยแล้วที่ {output_pdf}")
    
    return output_pdf
    
# ตัวอย่างการใช้งาน
if __name__ == "__main__":
    input_directory = "[path]"
    output_directory = "[path]"
    
    # ลดความละเอียดของรูปภาพเป็น 150 dpi และตั้งค่าคุณภาพ JPEG เป็น 85%
    results = find_and_compress_pdfs(input_directory, output_directory, dpi=150, image_quality=85)
    
    merge_pdfs(input_directory, output_directory)
    
    print("\nการทำงานเสร็จสิ้น!")