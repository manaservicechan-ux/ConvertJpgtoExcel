import streamlit as st
import pandas as pd
import easyocr
import re
from PIL import Image
import io

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Image Table to Excel Extractor", layout="centered")
st.title("📸 แปลงรูปภาพตารางเลข 6 หลัก เป็น Excel")
st.write("อัพโหลดรูปภาพตารางระบุเลข 6 หลัก (รวมหลายไฟล์ได้) เพื่อรวมเป็น Excel คอลัมน์เดียว")

# โหลดตัวอ่าน OCR (ใช้ภาษาอังกฤษเป็นหลักสำหรับตัวเลข)
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['en'], gpu=False)

reader = load_ocr()

# 1. ส่วนของการอัพโหลดไฟล์ (รองรับทั้งคอมและมือถือผ่านหน้าเว็บ)
uploaded_files = st.file_uploader(
    "เลือกรูปภาพตารางของคุณ (อัพโหลดได้หลายไฟล์พร้อมกัน)", 
    type=["png", "jpg", "jpeg"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.info(import_status := f"เลือกไฟล์ทั้งหมด {len(uploaded_files)} ไฟล์ กำลังเตรียมประมวลผล...")
    
    all_data = []  # เก็บข้อมูลทั้งหมด [เลข 6 หลัก, ชื่อไฟล์]
    
    # วนลูปประมวลผลทีละภาพ
    for uploaded_file in uploaded_files:
        with st.spinner(f"กำลังอ่านไฟล์: {uploaded_file.name} ..."):
            # อ่านรูปภาพ
            image = Image.open(uploaded_file)
            image_bytes = uploaded_file.getvalue()
            
            # ใช้ EasyOCR อ่านข้อความในภาพ
            results = reader.readtext(image_bytes)
            
            for (bbox, text, prob) in results:
                # ทำความสะอาดข้อความ ลบช่องว่างออก
                clean_text = text.replace(" ", "").strip()
                
                # ใช้ Regex ค้นหาตัวเลข 6 หลักติดกัน (เช่น 001234, 995678)
                match = re.search(r'^\d{6}$', clean_text)
                
                if match:
                    digit_6 = match.group(0)
                    all_data.append({
                        "ตัวเลข 6 หลัก": digit_6,
                        "หมายเหตุ (ชื่อไฟล์)": uploaded_file.name
                    })

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
        
        # 3. ปุ่มดาวน์โหลดไฟล์ Excel (กดเซฟลงเครื่องคอมหรือมือถือได้เลย)
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel",
            data=buffer,
            file_name="extracted_6_digits.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ ไม่พบตัวเลข 6 หลักในรูปภาพที่อัพโหลดเลย กรุณาตรวจสอบความชัดของภาพ")