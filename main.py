import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
from PIL import Image
from io import BytesIO
from datetime import datetime
import os
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

# --- ตั้งค่าหน้าเว็บ ---
st.set_page_config(page_title="Asset System", layout="wide")
st.title("📦 ระบบจัดการทรัพย์สิน (Online)")

# --- เชื่อมต่อ Google Sheets ---
@st.cache_resource
def connect_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scope)
    client = gspread.authorize(creds)
    return client.open("Your_Sheet_Name").worksheet("online")

# หมายเหตุ: สำหรับ st.secrets ให้ไปตั้งค่าใน Streamlit Cloud Dashboard 
# หรือถ้ายังไม่ถนัด ให้ใช้ไฟล์ credentials.json เหมือนเดิมได้ครับ

FIELDS = [
    "ID-Auto", "รูปภาพ", "QR-CODE", "บริษัท", "สถานะทรัพย์สิน", 
    "กลุ่มทรัพย์สิน", "รหัสทรัพย์สิน", "ชื่อทรัพย์สิน1", "แผนก", 
    "วันที่รับเข้าทะเบียน", "วันที่ตัดจากทะเบียน", "หน่วยนับ", 
    "จำนวน", "มูลค่าทุน", "ค่าเสื่อมสะสม", "มูลค่าคงเหลือ", "ข้อมูล ณ วันที่"
]

# --- โหลดข้อมูล ---
try:
    # ในตัวอย่างนี้ขอใช้แบบไฟล์เพื่อความง่ายตามเดิม
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open("Your_Sheet_Name").worksheet("online")
    all_data = sheet.get_all_values()
    df = pd.DataFrame(all_data[1:], columns=FIELDS)
except Exception as e:
    st.error(f"เชื่อมต่อผิดพลาด: {e}")
    df = pd.DataFrame(columns=FIELDS)

# --- ส่วน Filter (ด้านข้าง) ---
st.sidebar.header("🔍 ตัวกรองข้อมูล")
search_id = st.sidebar.text_input("ค้นหา ID-Auto / ชื่อ")
search_comp = st.sidebar.text_input("ค้นหา บริษัท")

# --- การกรองข้อมูล ---
filtered_df = df.copy()
if search_id:
    filtered_df = filtered_df[filtered_df['ID-Auto'].str.contains(search_id) | filtered_df['ชื่อทรัพย์สิน1'].str.contains(search_id)]
if search_comp:
    filtered_df = filtered_df[filtered_df['บริษัท'].str.contains(search_comp)]

# --- แสดงผลตาราง ---
st.write(f"พบข้อมูล {len(filtered_df)} รายการ")
selected_row = st.dataframe(filtered_df, use_container_width=True, on_select="rerun", selection_mode="single-row")

# --- ส่วนออกรายงาน PDF ---
if len(selected_row.selection.rows) > 0:
    idx = selected_row.selection.rows[0]
    data = filtered_df.iloc[idx]
    
    st.divider()
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.image(data['รูปภาพ'], caption="รูปทรัพย์สิน", width=250)
        st.image(data['QR-CODE'], caption="QR Code", width=150)
        
    with col2:
        st.subheader(f"รายละเอียด: {data['ชื่อทรัพย์สิน1']}")
        for f in FIELDS[3:]:
            st.write(f"**{f}:** {data[f]}")

    # ฟังก์ชันสร้าง PDF
    def generate_pdf(item):
        buf = BytesIO()
        c = canvas.Canvas(buf, pagesize=A4)
        width, height = A4
        
        # โหลดฟอนต์ (ตรวจสอบชื่อไฟล์ให้ตรงกับใน GitHub)
        try:
            pdfmetrics.registerFont(TTFont('ThaiBold', 'THSARABUN BOLD.ttf'))
            c.setFont('ThaiBold', 18)
        except:
            c.setFont('Helvetica-Bold', 18)

        # วาด QR (บนขวา)
        try:
            qr = ImageReader(item['QR-CODE'])
            c.drawImage(qr, width-130, height-130, 100, 100)
        except: pass

        c.drawString(50, height-50, f"รายงาน: {item['ชื่อทรัพย์สิน1']}")
        y = height - 100
        c.setFont('ThaiBold', 14)
        for f in FIELDS:
            if f not in ["รูปภาพ", "QR-CODE"]:
                c.drawString(70, y, f"• {f}: {item[f]}")
                y -= 25
        
        # วาดรูป (ล่าง)
        try:
            img = ImageReader(item['รูปภาพ'])
            c.drawImage(img, 70, 50, 250, 180, preserveAspectRatio=True)
        except: pass
        
        c.save()
        return buf.getvalue()

    pdf_data = generate_pdf(data)
    st.download_button(
        label="📥 ดาวน์โหลดเอกสาร PDF",
        data=pdf_data,
        file_name=f"Report_{data['ID-Auto']}.pdf",
        mime="application/pdf"
    )
