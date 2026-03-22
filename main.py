import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from PIL import Image
from io import BytesIO
import requests
import re
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
# เพิ่มไลบรารีสำหรับการสแกน QR Code
from streamlit_qrcode_scanner import qrcode_scanner 

# --- 1. ฟังก์ชันจัดการ URL และ QR Code ---
def get_drive_direct_link(cell_value):
    if not cell_value: return None
    url_match = re.search(r'https?://[^\s"]+', str(cell_value))
    if not url_match: return None
    url = url_match.group(0)
    if "drive.google.com" in url:
        file_id = ""
        if "/d/" in url: file_id = url.split("/d/")[1].split("/")[0].split("?")[0]
        elif "id=" in url: file_id = url.split("id=")[1].split("&")[0]
        if file_id: return f"https://drive.google.com/thumbnail?sz=w1000&id={file_id}"
    return url

def get_qr_url(id_text):
    if not id_text: return None
    return f"https://quickchart.io/qr?text={id_text}&size=150"

def download_image(url):
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        if resp.status_code == 200: return BytesIO(resp.content)
    except: return None
    return None

# --- 2. เชื่อมต่อ Google Sheets ---
st.set_page_config(page_title="ระบบจัดการทรัพย์สิน", layout="wide")
st.title("📦 ระบบจัดการทรัพย์สิน (Scan & Report)")

FIELDS = [
    "ID-Auto", "รูปภาพ", "QR-CODE", "บริษัท", "สถานะทรัพย์สิน", 
    "กลุ่มทรัพย์สิน", "รหัสทรัพย์สิน", "ชื่อทรัพย์สิน1", "แผนก", 
    "วันที่รับเข้าทะเบียน", "วันที่ตัดจากทะเบียน", "หน่วยนับ", 
    "จำนวน", "มูลค่าทุน", "ค่าเสื่อมสะสม", "มูลค่าคงเหลือ", "ข้อมูล ณ วันที่"
]

@st.cache_resource
def get_data_from_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_info = st.secrets["gcp_service_account"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
        client = gspread.authorize(creds)
        SHEET_ID = "1Pp2XffqRBtlyDu6NHDmFA6VcbdCmZPEv-1p3ETCSb5o" 
        sh = client.open_by_key(SHEET_ID) 
        worksheet = sh.worksheet("online")
        data = worksheet.get_all_values()
        if len(data) > 1: return pd.DataFrame(data[1:], columns=FIELDS)
        return pd.DataFrame(columns=FIELDS)
    except Exception as e:
        st.error(f"❌ เชื่อมต่อผิดพลาด: {e}")
        return None

df = get_data_from_sheets()

if df is not None:
    # --- 3. Sidebar: การค้นหาและสแกน QR ---
    st.sidebar.header("🔍 ค้นหาทรัพย์สิน")
    
    # ส่วนสแกน QR Code
    st.sidebar.subheader("📷 สแกน QR Code")
    qr_code_value = qrcode_scanner(key='scanner') # เปิดกล้องสแกน
    
    # ช่องค้นหา ID (ถ้าสแกนได้ ให้เอาค่าใส่ในช่องนี้อัตโนมัติ)
    if qr_code_value:
        st.sidebar.success(f"สแกนสำเร็จ: {qr_code_value}")
        search_id = st.sidebar.text_input("ค้นหา ID-Auto", value=qr_code_value)
    else:
        search_id = st.sidebar.text_input("ค้นหา ID-Auto", placeholder="พิมพ์ ID หรือสแกน...")

    search_comp = st.sidebar.text_input("ค้นหา บริษัท", placeholder="พิมพ์ชื่อบริษัท...")
    
    # กรองช่วงวันที่ (Start - End)
    st.sidebar.subheader("📅 กรองช่วงวันที่รับเข้า")
    date_range = st.sidebar.date_input(
        "เลือกช่วงวันที่",
        value=(datetime(2025, 1, 1), datetime.now()), # ค่าเริ่มต้น
        format="DD/MM/YYYY"
    )

    # --- 4. Logic การกรองข้อมูล ---
    filtered_df = df.copy()
    
    # กรอง ID และ บริษัท
    if search_id:
        filtered_df = filtered_df[filtered_df['ID-Auto'].str.contains(search_id, case=False, na=False)]
    if search_comp:
        filtered_df = filtered_df[filtered_df['บริษัท'].str.contains(search_comp, case=False, na=False)]
    
    # กรองช่วงวันที่ (แปลงคอลัมน์ใน Sheet เป็น datetime ก่อนเปรียบเทียบ)
    if len(date_range) == 2:
        start_date, end_date = date_range
        # พยายามแปลงวันที่จาก Sheet (สมมติว่าเป็นรูปแบบ วว/ดด/ปปปป)
        filtered_df['temp_date'] = pd.to_datetime(filtered_df['วันที่รับเข้าทะเบียน'], dayfirst=True, errors='coerce')
        filtered_df = filtered_df[
            (filtered_df['temp_date'].dt.date >= start_date) & 
            (filtered_df['temp_date'].dt.date <= end_date)
        ]

    # --- 5. แสดงผลตาราง ---
    st.write(f"📊 พบข้อมูลทั้งหมด **{len(filtered_df)}** รายการ")
    selection = st.dataframe(filtered_df.drop(columns=['temp_date'], errors='ignore'), use_container_width=True, on_select="rerun", selection_mode="single-row", hide_index=True)

    if len(selection.selection.rows) > 0:
        item = filtered_df.iloc[selection.selection.rows[0]]
        st.divider()
        
        col_img, col_txt = st.columns([1, 2])
        with col_img:
            img_url = get_drive_direct_link(item['รูปภาพ'])
            qr_url = get_qr_url(item['ID-Auto'])
            if img_url: st.image(img_url, caption="รูปทรัพย์สิน", width=300)
            if qr_url: st.image(qr_url, caption="QR-CODE", width=150)
        
        with col_txt:
            st.subheader(f"📄 {item['ชื่อทรัพย์สิน1']}")
            d1, d2 = st.columns(2)
            for i, f in enumerate(FIELDS):
                if f not in ["รูปภาพ", "QR-CODE"]:
                    target = d1 if i % 2 == 0 else d2
                    target.write(f"**{f}:** {item[f]}")

        # --- 6. สร้าง PDF (ใช้ฟังก์ชันเดิมที่เสถียรแล้ว) ---
        def generate_pdf(data):
            buf = BytesIO()
            c = canvas.Canvas(buf, pagesize=A4)
            w, h = A4
            try:
                pdfmetrics.registerFont(TTFont('ThaiBold', 'THSARABUN BOLD.ttf'))
                c.setFont('ThaiBold', 22)
            except: c.setFont('Helvetica-Bold', 18)
            
            c.drawString(50, h-60, "รายงานข้อมูลทรัพย์สิน")
            c.line(50, h-70, w-50, h-70)
            
            qr_pdf_url = get_qr_url(data['ID-Auto'])
            if qr_pdf_url:
                qr_data = download_image(qr_pdf_url)
                if qr_data: c.drawImage(ImageReader(qr_data), w-130, h-130, 80, 80)

            c.setFont('ThaiBold', 14)
            y = h - 100
            for f in FIELDS:
                if f not in ["รูปภาพ", "QR-CODE"]:
                    c.drawString(70, y, f"• {f}: {data[f]}")
                    y -= 22

            img_pdf_url = get_drive_direct_link(data['รูปภาพ'])
            if img_pdf_url:
                img_data = download_image(img_pdf_url)
                if img_data: c.drawImage(ImageReader(img_data), 70, 50, width=250, height=180, preserveAspectRatio=True)

            c.save()
            return buf.getvalue()

        st.download_button("📥 ดาวน์โหลดรายงาน PDF", data=generate_pdf(item), file_name=f"{item['ID-Auto']}.pdf")
