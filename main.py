import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from PIL import Image
from io import BytesIO
import requests
import re
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

# --- 1. ฟังก์ชันดึง URL และแปลงเป็น Direct Link (เน้นใช้ Thumbnail API) ---
def get_drive_direct_link(cell_value):
    if not cell_value:
        return None
    
    # แกะ URL ออกจากสูตร =IMAGE("...") หรือข้อความธรรมดา
    url_match = re.search(r'https?://[^\s"]+', str(cell_value))
    if not url_match:
        return None
    url = url_match.group(0)
    
    if "drive.google.com" in url:
        file_id = ""
        if "/d/" in url:
            file_id = url.split("/d/")[1].split("/")[0].split("?")[0]
        elif "id=" in url:
            file_id = url.split("id=")[1].split("&")[0]
        
        if file_id:
            # ใช้ thumbnail API จะเสถียรกว่าการใช้ uc?export=download
            return f"https://drive.google.com/thumbnail?sz=w1000&id={file_id}"
    
    return url

# --- 2. ฟังก์ชันดาวน์โหลดรูปภาพ ---
def download_image(url):
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        if resp.status_code == 200:
            return BytesIO(resp.content)
    except:
        return None
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
        if len(data) > 1:
            return pd.DataFrame(data[1:], columns=FIELDS)
        return pd.DataFrame(columns=FIELDS)
    except Exception as e:
        st.error(f"❌ เชื่อมต่อผิดพลาด: {e}")
        return None

# --- ส่วนของการสร้างหน้าเว็บและตาราง (ยึดตามโค้ดเดิมของคุณ) ---
st.set_page_config(page_title="ระบบจัดการทรัพย์สิน", layout="wide")
st.title("📦 ระบบจัดการทรัพย์สิน (Online Report)")

FIELDS = ["ID-Auto", "รูปภาพ", "QR-CODE", "บริษัท", "สถานะทรัพย์สิน", "กลุ่มทรัพย์สิน", "รหัสทรัพย์สิน", "ชื่อทรัพย์สิน1", "แผนก", "วันที่รับเข้าทะเบียน", "วันที่ตัดจากทะเบียน", "หน่วยนับ", "จำนวน", "มูลค่าทุน", "ค่าเสื่อมสะสม", "มูลค่าคงเหลือ", "ข้อมูล ณ วันที่"]

df = get_data_from_sheets()

if df is not None:
    # Sidebar
    st.sidebar.header("🔍 กรองข้อมูล")
    search_id = st.sidebar.text_input("ค้นหา ID-Auto")
    
    filtered_df = df.copy()
    if search_id:
        filtered_df = filtered_df[filtered_df['ID-Auto'].str.contains(search_id, na=False)]

    st.write(f"📊 พบข้อมูล {len(filtered_df)} รายการ")
    selection = st.dataframe(filtered_df, on_select="rerun", selection_mode="single-row", hide_index=True)

    if len(selection.selection.rows) > 0:
        item = filtered_df.iloc[selection.selection.rows[0]]
        st.divider()
        
        col_img, col_txt = st.columns([1, 2])
        with col_img:
            img_url = get_drive_direct_link(item['รูปภาพ'])
            qr_url = get_drive_direct_link(item['QR-CODE'])
            if img_url: st.image(img_url, caption="รูปทรัพย์สิน", width=300)
            if qr_url: st.image(qr_url, caption="QR-CODE", width=150)
        
        with col_txt:
            st.subheader(f"📄 {item['ชื่อทรัพย์สิน1']}")
            for f in FIELDS:
                if f not in ["รูปภาพ", "QR-CODE"]:
                    st.write(f"**{f}:** {item[f]}")

        # --- 4. ฟังก์ชันสร้าง PDF ---
        def generate_pdf(data):
            buf = BytesIO()
            c = canvas.Canvas(buf, pagesize=A4)
            w, h = A4
            
            try:
                pdfmetrics.registerFont(TTFont('ThaiBold', 'THSARABUN BOLD.ttf'))
                c.setFont('ThaiBold', 22)
            except:
                c.setFont('Helvetica-Bold', 18)

            c.drawString(50, h-60, "รายงานข้อมูลทรัพย์สิน")
            c.line(50, h-70, w-50, h-70)

            # QR Code (มุมขวาบน)
            q_link = get_drive_direct_link(data['QR-CODE'])
            if q_link:
                q_data = download_image(q_link)
                if q_data:
                    c.drawImage(ImageReader(q_data), w-130, h-130, 80, 80)

            # ข้อมูลตัวอักษร
            c.setFont('ThaiBold', 14)
            y_pos = h - 100
            for f in FIELDS:
                if f not in ["รูปภาพ", "QR-CODE"]:
                    c.drawString(70, y_pos, f"• {f}: {data[f]}")
                    y_pos -= 22

            # รูปหลัก (ด้านล่าง)
            i_link = get_drive_direct_link(data['รูปภาพ'])
            if i_link:
                i_data = download_image(i_link)
                if i_data:
                    c.drawImage(ImageReader(i_data), 70, 50, width=250, height=180, preserveAspectRatio=True)

            c.save()
            return buf.getvalue()

        st.download_button("📥 ดาวน์โหลด PDF", data=generate_pdf(item), file_name=f"{item['ID-Auto']}.pdf")
