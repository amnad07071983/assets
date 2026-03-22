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

# --- 1. ตั้งค่าหน้าจอแอป ---
st.set_page_config(page_title="ระบบจัดการทรัพย์สิน", layout="wide")
st.title("📦 ระบบจัดการทรัพย์สิน (Online Report)")

# รายชื่อฟิลด์ทั้ง 17 ฟิลด์
FIELDS = [
    "ID-Auto", "รูปภาพ", "QR-CODE", "บริษัท", "สถานะทรัพย์สิน", 
    "กลุ่มทรัพย์สิน", "รหัสทรัพย์สิน", "ชื่อทรัพย์สิน1", "แผนก", 
    "วันที่รับเข้าทะเบียน", "วันที่ตัดจากทะเบียน", "หน่วยนับ", 
    "จำนวน", "มูลค่าทุน", "ค่าเสื่อมสะสม", "มูลค่าคงเหลือ", "ข้อมูล ณ วันที่"
]

# --- 2. ฟังก์ชันจัดการ URL รูปภาพ (แกะสูตรและแปลงเป็น Direct Link) ---
def get_drive_direct_link(cell_value):
    if not cell_value:
        return None
    # แกะ URL ออกจากสูตร =IMAGE("...") หรือข้อความธรรมดา
    url_match = re.search(r'https?://[^\s"]+', cell_value)
    if not url_match:
        return None
    url = url_match.group(0)
    # ถ้าเป็น Google Drive ให้แปลงเป็นลิงก์ดาวน์โหลดตรง
    if "drive.google.com" in url:
        file_id = ""
        if "/d/" in url:
            file_id = url.split("/d/")[1].split("/")[0].split("?")[0]
        elif "id=" in url:
            file_id = url.split("id=")[1].split("&")[0]
        if file_id:
            return f"https://drive.google.com/uc?export=download&id={file_id}"
    return url

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

# --- 4. เริ่มโหลดข้อมูลและแสดงผล ---
df = get_data_from_sheets()

if df is not None:
    # ส่วน Sidebar สำหรับกรองข้อมูล
    st.sidebar.header("🔍 กรองข้อมูล")
    search_id = st.sidebar.text_input("ค้นหา ID-Auto")
    search_comp = st.sidebar.text_input("ค้นหา บริษัท")

    filtered_df = df.copy()
    if search_id:
        filtered_df = filtered_df[filtered_df['ID-Auto'].str.contains(search_id, na=False)]
    if search_comp:
        filtered_df = filtered_df[filtered_df['บริษัท'].str.contains(search_comp, na=False)]

    st.write(f"📊 พบข้อมูล {len(filtered_df)} รายการ")
    selection = st.dataframe(filtered_df, on_select="rerun", selection_mode="single-row", hide_index=True)

    if len(selection.selection.rows) > 0:
        item = filtered_df.iloc[selection.selection.rows[0]]
        st.divider()
        
        # แสดงผลหน้าเว็บ
        col_img, col_txt = st.columns([1, 2])
        with col_img:
            # แปลงลิงก์เพื่อแสดงผลบน Streamlit
            st.image(get_drive_direct_link(item['รูปภาพ']) or "https://via.placeholder.com/150", width=300)
            st.image(get_drive_direct_link(item['QR-CODE']) or "https://via.placeholder.com/100", width=100)
        
        with col_txt:
            st.subheader(f"📄 {item['ชื่อทรัพย์สิน1']}")
            for f in FIELDS:
                if f not in ["รูปภาพ", "QR-CODE"]:
                    st.write(f"**{f}:** {item[f]}")

        # --- 5. ฟังก์ชันสร้าง PDF ---
        def generate_pdf(data):
            buf = BytesIO()
            c = canvas.Canvas(buf, pagesize=A4)
            w, h = A4
            
            try:
                pdfmetrics.registerFont(TTFont('ThaiBold', 'THSARABUN BOLD.ttf'))
                c.setFont('ThaiBold', 22)
            except:
                c.setFont('Helvetica-Bold', 18)

            # วาดหัวข้อ
            c.drawString(50, h-60, "รายงานข้อมูลทรัพย์สิน")
            c.line(50, h-70, w-50, h-70)

            # วาด QR-CODE
            qr_url = get_drive_direct_link(data['QR-CODE'])
            if qr_url:
                qr_img = download_image(qr_url)
                if qr_img:
                    c.drawImage(ImageReader(qr_img), w-130, h-130, 80, 80)

            # รายละเอียดข้อมูล
            c.setFont('ThaiBold', 14)
            y = h - 100
            for f in FIELDS:
                if f not in ["รูปภาพ", "QR-CODE"]:
                    c.drawString(70, y, f"• {f}: {data[f]}")
                    y -= 22

            # วาดรูปทรัพย์สิน
            img_url = get_drive_direct_link(data['รูปภาพ'])
            if img_url:
                asset_img = download_image(img_url)
                if asset_img:
                    c.drawImage(ImageReader(asset_img), 70, 50, width=250, height=180, preserveAspectRatio=True)

            c.save()
            return buf.getvalue()

        st.download_button("📥 ดาวน์โหลด PDF", data=generate_pdf(item), file_name=f"{item['ID-Auto']}.pdf")
