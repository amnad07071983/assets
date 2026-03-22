import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from PIL import Image
from io import BytesIO
from datetime import datetime
import requests
import re
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
# เพิ่มส่วนสแกน QR
from streamlit_qrcode_scanner import qrcode_scanner 

# --- 1. ตั้งค่าหน้าจอแอป ---
st.set_page_config(page_title="ระบบจัดการทรัพย์สิน", layout="wide")
st.title("📦 ระบบจัดการทรัพย์สิน (Online Report)")

FIELDS = [
    "ID-Auto", "รูปภาพ", "QR-CODE", "บริษัท", "สถานะทรัพย์สิน", 
    "กลุ่มทรัพย์สิน", "รหัสทรัพย์สิน", "ชื่อทรัพย์สิน1", "แผนก", 
    "วันที่รับเข้าทะเบียน", "วันที่ตัดจากทะเบียน", "หน่วยนับ", 
    "จำนวน", "มูลค่าทุน", "ค่าเสื่อมสะสม", "มูลค่าคงเหลือ", "ข้อมูล ณ วันที่"
]

# --- 2. ฟังก์ชันเสริมสำหรับการจัดการรูปและ QR ---
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

# --- 3. ฟังก์ชันเชื่อมต่อ Google Sheets ---
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
        st.error(f"❌ ไม่สามารถเชื่อมต่อ Google Sheets ได้: {e}")
        return None

df = get_data_from_sheets()

if df is not None:
    # --- 4. ส่วนตัวกรองข้อมูล (Sidebar) ---
    st.sidebar.header("🔍 ค้นหาและกรองข้อมูล")
    
    # เพิ่มส่วนสแกน QR Code
    st.sidebar.subheader("📷 สแกน QR Code")
    scanned_value = qrcode_scanner(key='scanner')
    
    # ค้นหา ID (รองรับทั้งพิมพ์และสแกน)
    if scanned_value:
        st.sidebar.success(f"สแกนพบ: {scanned_value}")
        search_id = st.sidebar.text_input("ค้นหา ID-Auto", value=scanned_value)
    else:
        search_id = st.sidebar.text_input("ค้นหา ID-Auto", placeholder="พิมพ์ ID หรือสแกน...")

    search_comp = st.sidebar.text_input("ค้นหา บริษัท", placeholder="พิมพ์ชื่อบริษัท...")
    
    # เพิ่มการกรองช่วงวันที่ (Start - End)
    st.sidebar.subheader("📅 ช่วงวันที่รับเข้าทะเบียน")
    date_selection = st.sidebar.date_input(
        "เลือกช่วงวันที่",
        value=(datetime(1111, 1, 1), datetime.now()),
        format="DD/MM/YYYY"
    )

    # ตรรกะการกรอง
    filtered_df = df.copy()
    if search_id:
        filtered_df = filtered_df[filtered_df['ID-Auto'].str.contains(search_id, case=False, na=False)]
    if search_comp:
        filtered_df = filtered_df[filtered_df['บริษัท'].str.contains(search_comp, case=False, na=False)]
    
    # กรองวันที่จากช่วงที่เลือก
    if isinstance(date_selection, tuple) and len(date_selection) == 2:
        start_d, end_d = date_selection
        filtered_df['dt_temp'] = pd.to_datetime(filtered_df['วันที่รับเข้าทะเบียน'], dayfirst=True, errors='coerce')
        filtered_df = filtered_df[
            (filtered_df['dt_temp'].dt.date >= start_d) & 
            (filtered_df['dt_temp'].dt.date <= end_d)
        ]

    # --- 5. แสดงตารางรายการทรัพย์สิน ---
    st.write(f"📊 พบข้อมูลทั้งหมด **{len(filtered_df)}** รายการ")
    selection = st.dataframe(
        filtered_df.drop(columns=['dt_temp'], errors='ignore'), 
        use_container_width=True, 
        on_select="rerun", 
        selection_mode="single-row",
        hide_index=True
    )

    # --- 6. แสดงรายละเอียดและการดาวน์โหลด PDF ---
    if len(selection.selection.rows) > 0:
        item = filtered_df.iloc[selection.selection.rows[0]]
        st.divider()
        
        col_img, col_detail = st.columns([1, 2])
        with col_img:
            img_url = get_drive_direct_link(item['รูปภาพ'])
            qr_url = get_qr_url(item['ID-Auto'])
            if img_url: st.image(img_url, caption="📸 รูปทรัพย์สิน", use_container_width=True)
            if qr_url: st.image(qr_url, caption="🔗 QR-CODE (ID-Auto)", width=150)
            
        with col_detail:
            st.subheader(f"📄 รายละเอียด: {item['ชื่อทรัพย์สิน1']}")
            info_1, info_2 = st.columns(2)
            for i, field in enumerate(FIELDS):
                if field not in ["รูปภาพ", "QR-CODE"]:
                    target = info_1 if i % 2 == 0 else info_2
                    target.write(f"**{field}:** {item[field]}")

            # --- 7. ฟังก์ชันสร้าง PDF ---
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
                
                # QR Code ใน PDF
                q_url = get_qr_url(data['ID-Auto'])
                if q_url:
                    q_data = download_image(q_url)
                    if q_data: c.drawImage(ImageReader(q_data), w-130, h-130, 80, 80)

                c.setFont('ThaiBold', 14)
                curr_y = h - 110
                for f in FIELDS:
                    if f not in ["รูปภาพ", "QR-CODE"]:
                        c.drawString(70, curr_y, f"• {f}: {data[f]}")
                        curr_y -= 22
                
                # รูปหลักใน PDF
                i_url = get_drive_direct_link(data['รูปภาพ'])
                if i_url:
                    i_data = download_image(i_url)
                    if i_data: c.drawImage(ImageReader(i_data), 70, 50, width=250, height=180, preserveAspectRatio=True)
                
                c.save()
                return buf.getvalue()

            st.download_button(
                label="📥 ดาวน์โหลดรายงาน PDF",
                data=generate_pdf(item),
                file_name=f"Report_{item['ID-Auto']}.pdf",
                mime="application/pdf"
            )
else:
    st.info("💡 กรุณาตรวจสอบการเชื่อมต่อ Google Sheets")
