import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from PIL import Image
from io import BytesIO
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

# --- 1. ตั้งค่าหน้าจอแอป ---
st.set_page_config(page_title="ระบบจัดการทรัพย์สิน", layout="wide")
st.title("📦 ระบบจัดการทรัพย์สิน (Online Report)")

# รายชื่อฟิลด์ทั้ง 17 ฟิลด์ตามโครงสร้าง Sheet "online"
FIELDS = [
    "ID-Auto", "รูปภาพ", "QR-CODE", "บริษัท", "สถานะทรัพย์สิน", 
    "กลุ่มทรัพย์สิน", "รหัสทรัพย์สิน", "ชื่อทรัพย์สิน1", "แผนก", 
    "วันที่รับเข้าทะเบียน", "วันที่ตัดจากทะเบียน", "หน่วยนับ", 
    "จำนวน", "มูลค่าทุน", "ค่าเสื่อมสะสม", "มูลค่าคงเหลือ", "ข้อมูล ณ วันที่"
]

# --- 2. ฟังก์ชันเชื่อมต่อ Google Sheets (ใช้ ID และ Secrets) ---
@st.cache_resource
def get_data_from_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        # ดึงข้อมูลจาก Secrets ที่ตั้งค่าไว้ใน Streamlit Cloud
        creds_info = st.secrets["gcp_service_account"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
        client = gspread.authorize(creds)
        
        # 📍 แก้ไข: ใส่ ID ของ Google Sheet ของคุณที่นี่ 📍
        # (ID คือตัวอักษรยาวๆ ใน URL ของ Sheet ระหว่าง /d/ และ /edit)
        SHEET_ID = "1Pp2XffqRBtlyDu6NHDmFA6VcbdCmZPEv-1p3ETCSb5o" 
        
        sh = client.open_by_key(SHEET_ID) 
        worksheet = sh.worksheet("online")
        
        data = worksheet.get_all_values()
        if len(data) > 1:
            # สร้าง DataFrame โดยข้ามแถวหัวตาราง (แถวที่ 1)
            return pd.DataFrame(data[1:], columns=FIELDS)
        return pd.DataFrame(columns=FIELDS)
    except Exception as e:
        st.error(f"❌ ไม่สามารถเชื่อมต่อ Google Sheets ได้: {e}")
        return None

# --- 3. โหลดข้อมูล ---
df = get_data_from_sheets()

if df is not None:
    # --- 4. ส่วนตัวกรองข้อมูล (Sidebar) ---
    st.sidebar.header("🔍 ค้นหาและกรองข้อมูล")
    
    # กรองด้วยข้อความ
    search_id = st.sidebar.text_input("ค้นหา ID-Auto", placeholder="พิมพ์ ID...")
    search_comp = st.sidebar.text_input("ค้นหา บริษัท", placeholder="พิมพ์ชื่อบริษัท...")
    
    # กรองด้วยวันที่ (รับเข้า / ตัดจากทะเบียน)
    st.sidebar.subheader("📅 กรองด้วยวันที่")
    date_in = st.sidebar.text_input("วันที่รับเข้า (วว/ดด/ปปปป)", placeholder="เช่น 01/01/2024")
    date_out = st.sidebar.text_input("วันที่ตัดทะเบียน (วว/ดด/ปปปป)", placeholder="เช่น 31/12/2024")

    # ตรรกะการกรอง (Filtering)
    filtered_df = df.copy()
    if search_id:
        filtered_df = filtered_df[filtered_df['ID-Auto'].str.contains(search_id, case=False, na=False)]
    if search_comp:
        filtered_df = filtered_df[filtered_df['บริษัท'].str.contains(search_comp, case=False, na=False)]
    if date_in:
        filtered_df = filtered_df[filtered_df['วันที่รับเข้าทะเบียน'].str.contains(date_in, na=False)]
    if date_out:
        filtered_df = filtered_df[filtered_df['วันที่ตัดจากทะเบียน'].str.contains(date_out, na=False)]

    # --- 5. แสดงตารางรายการทรัพย์สิน ---
    st.write(f"📊 พบข้อมูลทั้งหมด **{len(filtered_df)}** รายการ")
    
    # สร้างตารางที่คลิกเลือกแถวได้
    selection = st.dataframe(
        filtered_df, 
        use_container_width=True, 
        on_select="rerun", 
        selection_mode="single-row",
        hide_index=True
    )

    # --- 6. แสดงรายละเอียดและการดาวน์โหลด PDF เมื่อเลือกแถว ---
    if len(selection.selection.rows) > 0:
        row_idx = selection.selection.rows[0]
        item = filtered_df.iloc[row_idx]
        
        st.divider()
        col_img, col_detail = st.columns([1, 2])
        
        with col_img:
            # แสดงรูปภาพจากลิงก์ใน Sheet
            if item['รูปภาพ']:
                st.image(item['รูปภาพ'], caption="📸 รูปทรัพย์สิน", use_container_width=True)
            if item['QR-CODE']:
                st.image(item['QR-CODE'], caption="🔗 QR-CODE", width=150)
            
        with col_detail:
            st.subheader(f"📄 รายละเอียดทรัพย์สิน: {item['ชื่อทรัพย์สิน1']}")
            
            # แบ่งการแสดงผลข้อมูลเป็น 2 คอลัมน์ย่อย
            info_1, info_2 = st.columns(2)
            for i, field in enumerate(FIELDS):
                if field not in ["รูปภาพ", "QR-CODE"]:
                    target = info_1 if i % 2 == 0 else info_2
                    target.write(f"**{field}:** {item[field]}")

            # --- 7. ฟังก์ชันสร้าง PDF (ฟอนต์ไทย + รูป + QR) ---
            def generate_pdf(data):
                buf = BytesIO()
                c = canvas.Canvas(buf, pagesize=A4)
                w, h = A4
                
                # ลงทะเบียนฟอนต์ (ต้องมีไฟล์ THSARABUN BOLD.ttf ใน GitHub)
                try:
                    pdfmetrics.registerFont(TTFont('ThaiBold', 'THSARABUN BOLD.ttf'))
                    c.setFont('ThaiBold', 22)
                except:
                    c.setFont('Helvetica-Bold', 18)

                # วาด QR-CODE (มุมบนขวา)
                try:
                    qr_reader = ImageReader(data['QR-CODE'])
                    c.drawImage(qr_reader, w-130, h-130, 100, 100)
                except: pass

                # หัวข้อเอกสาร
                c.drawString(50, h-60, "รายงานข้อมูลทรัพย์สิน")
                c.setLineWidth(1)
                c.line(50, h-70, w-150, h-70)
                
                # รายละเอียดข้อมูล (เรียงลงมา)
                c.setFont('ThaiBold', 15)
                curr_y = h - 110
                for f in FIELDS:
                    if f not in ["รูปภาพ", "QR-CODE"]:
                        c.drawString(70, curr_y, f"• {f}: {data[f]}")
                        curr_y -= 24
                
                # วาดรูปทรัพย์สิน (ด้านล่าง)
                try:
                    asset_img = ImageReader(data['รูปภาพ'])
                    # วาดรูปที่ความสูง 50 จากขอบล่าง
                    c.drawImage(asset_img, 70, 50, width=280, height=200, preserveAspectRatio=True)
                except: pass
                
                c.save()
                return buf.getvalue()

            # ปุ่มดาวน์โหลด
            pdf_file = generate_pdf(item)
            st.download_button(
                label="📥 ดาวน์โหลดรายงาน PDF",
                data=pdf_file,
                file_name=f"Report_{item['ID-Auto']}.pdf",
                mime="application/pdf"
            )

else:
    st.info("💡 กรุณานำ Google Sheet ID มาใส่ในโค้ด และตรวจสอบการตั้งค่า Secrets ให้เรียบร้อย")
