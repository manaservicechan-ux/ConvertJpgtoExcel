import streamlit as st
import pandas as pd
import easyocr
import re
from PIL import Image
import io

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Image Table to Excel Extractor", layout="centered")
st.title("📸 แปลงรูปภาพตารางเลข 6 หลัก เป็น Excel")
st.write("เวอร์ชันรองรับมือถือ (ลดการใช้ RAM ป้องกันระบบล่ม)")

# โหลดตัวอ่าน OCR (บังคับใช้ CPU และปิดฟีเจอร์ที่กินแรมเยอะ)
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['en'], gpu=False)

try:
    reader = load_ocr()
except Exception as e:
    st.error(f"เกิดข้อผิดพลาดในการโหลดระบบอ่านข้อความ: {e}")

# 1. ส่วนของการอัพโหลดไฟล์ (รองรับทั้งคอมและมือถือผ่านหน้าเว็บ)
uploaded_files = st.file_uploader(
    "เลือกรูปภาพตารางของคุณ (อัพโหลดได้หลายไฟล์พร้อมกัน)", 
    type=["png", "jpg", "jpeg"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"เลือกไฟล์ทั้งหมด {len(uploaded_files)} ไฟล์ กำลังเตรียมประมวลผล...")
    
    all_data = []  # เก็บข้อมูลทั้งหมด [เลข 6 หลัก, ชื่อไฟล์]
    
    # วนลูปประมวลผลทีละภาพ
    for uploaded_file in uploaded_files:
        with st.spinner(f"กำลังอ่านไฟล์: {uploaded_file.name} ..."):
            try:
                # ─── ขั้นตอนสำคัญ: ย่อขนาดรูปภาพจากมือถือเพื่อประหยัด RAM ───
                image = Image.open(uploaded_file)
                
                # ถ้ารูปภาพกว้างเกิน 1200 พิกเซล ให้ย่อลงมา (ตัวเลขตารางยังชัดอยู่ แต่ไฟล์จะเล็กลงมาก)
                max_width = 1200
                if image.width > max_width:
                    ratio = max_width / float(image.width)
                    height = int((float(image.height) * float(ratio)))
                    image = image.resize((max_width, height), Image.Resampling.LANCZOS)
                
                # แปลงรูปภาพที่ย่อแล้วเป็น Bytes เพื่อส่งให้ EasyOCR
                img_byte_arr = io.BytesIO()
                image.save(img_byte_arr, format='JPEG', quality=85) # บีบอัดคุณภาพเหลือ 85% เพื่อลดขนาด
                image_bytes = img_byte_arr.getvalue()
                
                # ใช้ EasyOCR อ่านข้อความจากรูปที่ย่อแล้ว
                results = reader.readtext(image_bytes, detail=0) # detail=0 จะคืนค่าเฉพาะข้อความ ช่วยประหยัดแรมเพิ่มขึ้น
                
                for text in results:
                    # ทำความสะอาดข้อความ ลบช่องว่างออก
                    clean_text = text.replace(" ", "").strip()
                    
                    # ใช้ Regex ค้นหาตัวเลข 6 หลักติดกัน (00-99 นำหน้าได้)
                    match = re.search(r'^\d{6}$', clean_text)
                    
                    if match:
                        digit_6 = match.group(0)
                        all_data.append({
                            "ตัวเลข 6 หลัก": digit_6,
                            "หมายเหตุ (ชื่อไฟล์)": uploaded_file.name
                        })
            except Exception as img_err:
                st.error(f"ไม่สามารถประมวลผลไฟล์ {uploaded_file.name} ได้: {img_err}")

    # 2. นำข้อมูลมาสร้างเป็น DataFrame และตารางแสดงผล
    if all_data:
        df = pd.DataFrame(all_data)
        
        st.success("🎉 ประมวลผลเสร็จสิ้น!")
        st.write("### ตัวอย่างข้อมูลที่ดึงได้:")
        st.dataframe(df, use_container_width=True)
        
        # แปลง DataFrame เป็น Excel (เก็บใน Memory)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="รวมเลข 6 หลัก")
        
        buffer.seek(0)
        
        # 3. ปุ่มดาวน์โหลดไฟล์ Excel
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel",
            data=buffer,
            file_name="extracted_6_digits.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ ไม่พบตัวเลข 6 หลักในรูปภาพที่อัพโหลดเลย กรุณาตรวจสอบความชัดของภาพหรือลองถ่ายให้ใกล้ขึ้น")
