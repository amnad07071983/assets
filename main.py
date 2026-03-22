import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from PIL import Image
from io import BytesIO
from datetime import datetime
import os

# สำหรับสร้าง PDF
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

# --- 1. ตั้งค่าหน้าเว็บ ---
st.set_page_config(page_title="Asset Management System", layout="wide")
st.title("📦 ระบบจัดการทรัพย์สิน (Online Report)")

# --- 2. ฟิลด์ข้อมูล 17 ฟิลด์ ---
FIELDS = [
    "ID-Auto", "รูปภาพ", "QR-CODE", "บริษัท", "สถานะทรัพย์สิน", 
    "กลุ่มทรัพย์สิน", "รหัสทรัพย์สิน", "ชื่อทรัพย์สิน1", "แผนก", 
    "วันที่รับเข้าทะเบียน", "วันที่ตัดจากทะเบียน", "หน่วยนับ", 
    "จำนวน", "มูลค่าทุน", "ค่าเสื่อมสะสม", "มูลค่าคงเหลือ", "ข้อมูล ณ วันที่"
]

# --- 3. เชื่อมต่อ Google Sheets ผ่าน Secrets ---
@st.cache_resource
def connect_sheet():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        # ดึงข้อมูลจากช่อง Secrets ที่คุณกรอกไว้
        creds_info = st.secrets["gcp_service_account"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
        client = gspread.authorize(creds)
        # เปลี่ยนชื่อ "transport-appm1" เป็นชื่อไฟล์ Google Sheet ของคุณ
        return client.open("transport-appm1").worksheet("online")
    except Exception as e:
        st.error(f"❌ ไม่สามารถเชื่อมต่อ Google Sheets ได้: {e}")
        return None

# --- 4. ฟังก์ชันจัดการวันที่ ---
def parse_date(date_str):
    if not date_str: return None
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(str(date_str).strip(), fmt)
        except:
            continue
    return None

# --- 5. เริ่มทำงาน ---
sheet = connect_sheet()

if sheet:
    # โหลดข้อมูล
    raw_data = sheet.get_all_values()
    if len(raw_data) > 1:
        df = pd.DataFrame(raw_data[1:], columns=FIELDS)
        
        # --- ส่วน Filter (Sidebar) ---
        st.sidebar.header("🔍 ตัวกรองข้อมูล")
        search_keyword = st.sidebar.text_input("ค้นหา (ID / ชื่อทรัพย์สิน / บริษัท)")
        
        st.sidebar.subheader("📅 ช่วงวันที่รับเข้าทะเบียน")
        d_start = st.sidebar.date_value = st.sidebar.text_input("เริ่ม (วว/ดด/ปปปป)", placeholder="01/01/2024")
        d_end = st.sidebar.text_input("ถึง (วว/ดด/ปปปป)", placeholder="31/12/2024")

        # --- การกรองข้อมูล ---
        mask = pd.Series([True] * len(df))
        if search_keyword:
            mask &= (df['ID-Auto'].str.contains(search_keyword, case=False) | 
                     df['ชื่อทรัพย์สิน1'].str.contains(search_keyword, case=False) |
                     df['บริษัท'].str.contains(search_keyword, case=False))
        
        # กรองวันที่ (ถ้ามีการระบุ)
        date_s = parse_date(d_start)
        date_e = parse_date(d_end)
        if date_s or date_e:
            row_dates = df['วันที่รับเข้าทะเบียน'].apply(parse_date)
            if date_s: mask &= (row_dates >= date_s)
            if date_e: mask &= (row_dates <= date_e)

        filtered_df = df[mask]

        # --- แสดงตาราง ---
        st.subheader(f"📊 รายการทรัพย์สิน ({len(filtered_df)} รายการ)")
        # ใช้ Selection Mode เพื่อให้คลิกเลือกแถวที่จะพิมพ์ได้
        event = st.dataframe(
            filtered_df[['ID-Auto', 'ชื่อทรัพย์สิน1', 'บริษัท', 'วันที่รับเข้าทะเบียน', 'สถานะทรัพย์สิน']], 
            use_container_width=True, 
            on_select="rerun", 
            selection_mode="single-row"
        )

        # --- ส่วนแสดงรายละเอียดและปุ่มพิมพ์ PDF ---
        if len(event.selection.rows) > 0:
            idx = event.selection.rows[0]
            item = filtered_df.iloc[idx]
            
            st.divider()
            c1, c2 = st.columns([1, 2])
            
            with c1:
                st.image(item['รูปภาพ'], caption="รูปภาพทรัพย์สิน", use_container_width=True)
                st.image(item['QR-CODE'], caption="QR-CODE", width=150)
            
            with c2:
                st.subheader(f"📄 รายละเอียด: {item['ชื่อทรัพย์สิน1']}")
                # แสดงข้อมูลแบบ Grid
                info_cols = st.columns(2)
                for i, f in enumerate(FIELDS):
                    if f not in ["รูปภาพ", "QR-CODE"]:
                        info_cols[i % 2].write(f"**{f}:** {item[f]}")

                # ฟังก์ชันสร้าง PDF
                def make_pdf(data):
                    buffer = BytesIO()
                    c = canvas.Canvas(buffer, pagesize=A4)
                    w, h = A4
                    
                    # โหลดฟอนต์ภาษาไทยจากไฟล์ที่คุณมีใน GitHub
                    try:
                        pdfmetrics.registerFont(TTFont('ThaiBold', 'THSARABUN BOLD.ttf'))
                        c.setFont('ThaiBold', 18)
                    except:
                        c.setFont('Helvetica-Bold', 18)

                    # 1. QR-CODE บนขวา
                    try:
                        qr_img = ImageReader(data['QR-CODE'])
                        c.drawImage(qr_img, w-130, h-130, 100, 100)
                    except: pass

                    # 2. หัวข้อและข้อมูล
                    c.drawString(50, h-50, f"รายงานข้อมูลทรัพย์สิน")
                    c.setFont('ThaiBold', 14)
                    y_pos = h - 100
                    for f in FIELDS:
                        if f not in ["รูปภาพ", "QR-CODE"]:
                            c.drawString(70, y_pos, f"{f}: {data[f]}")
                            y_pos -= 22
                    
                    # 3. รูปภาพด้านล่าง
                    try:
                        main_img = ImageReader(data['รูปภาพ'])
                        c.drawImage(main_img, 70, y_pos - 180, width=250, height=180, preserveAspectRatio=True)
                    except: pass
                    
                    c.save()
                    return buffer.getvalue()

                # ปุ่มดาวน์โหลด
                pdf_bytes = make_pdf(item)
                st.download_button(
                    label="📥 ดาวน์โหลดเอกสาร PDF (TH Sarabun)",
                    data=pdf_bytes,
                    file_name=f"Asset_{item['ID-Auto']}.pdf",
                    mime="application/pdf"
                )
    else:
        st.warning("⚠️ ไม่พบข้อมูลในแผ่นงาน 'online'")
