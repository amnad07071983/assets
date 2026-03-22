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

# --- 1. ฟังก์ชันจัดการ URL, QR Code และการดาวน์โหลด ---
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

# --- 2. ตั้งค่าหน้าจอและเชื่อมต่อข้อมูล ---
st.set_page_config(page_title="ระบบจัดการทรัพย์สิน", layout="wide")
st.title("📦 ระบบจัดการทรัพย์สิน (Online Report)")

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

# --- 3. ส่วนการโหลดข้อมูลและการกรอง (Filter) ---
df = get_data_from_sheets()

if df is not None:
    # --- ส่วน Sidebar (ใส่ส่วนค้นหาบริษัทและวันที่กลับมา) ---
    st.sidebar.header("🔍 ค้นหาและกรองข้อมูล")
    search_id = st.sidebar.text_input("ค้นหา ID-Auto", placeholder="พิมพ์ ID...")
    search_comp = st.sidebar.text_input("ค้นหา บริษัท", placeholder="พิมพ์ชื่อบริษัท...")
    
    st.sidebar.subheader("📅 กรองด้วยวันที่")
    date_in = st.sidebar.text_input("วันที่รับเข้า (วว/ดด/ปปปป)", placeholder="เช่น 01/01/2025")
    date_out = st.sidebar.text_input("วันที่ตัดทะเบียน (วว/ดด/ปปปป)", placeholder="เช่น 31/12/2026")

    # ตรรกะการกรองข้อมูล
    filtered_df = df.copy()
    if search_id:
        filtered_df = filtered_df[filtered_df['ID-Auto'].str.contains(search_id, case=False, na=False)]
    if search_comp:
        filtered_df = filtered_df[filtered_df['บริษัท'].str.contains(search_comp, case=False, na=False)]
    if date_in:
        filtered_df = filtered_df[filtered_df['วันที่รับเข้าทะเบียน'].str.contains(date_in, na=False)]
    if date_out:
        filtered_df = filtered_df[filtered_df['วันที่ตัดจากทะเบียน'].str.contains(date_out, na=False)]

    # --- 4. แสดงตารางและการเลือกข้อมูล ---
    st.write(f"📊 พบข้อมูลทั้งหมด **{len(filtered_df)}** รายการ")
    selection = st.dataframe(
        filtered_df, 
        use_container_width=True, 
        on_select="rerun", 
        selection_mode="single-row",
        hide_index=True
    )

    if len(selection.selection.rows) > 0:
        item = filtered_df.iloc[selection.selection.rows[0]]
        st.divider()
        
        col_img, col_txt = st.columns([1, 2])
        with col_img:
            img_display_url = get_drive_direct_link(item['รูปภาพ'])
            qr_display_url = get_qr_url(item['ID-Auto'])
            if img_display_url: st.image(img_display_url, caption="รูปทรัพย์สิน", width=300)
            if qr_display_url: st.image(qr_display_url, caption="QR-CODE (สร้างจาก ID)", width=150)
        
        with col_txt:
            st.subheader(f"📄 รายละเอียด: {item['ชื่อทรัพย์สิน1']}")
            d_col1, d_col2 = st.columns(2)
            for i, field in enumerate(FIELDS):
                if field not in ["รูปภาพ", "QR-CODE"]:
                    target = d_col1 if i % 2 == 0 else d_col2
                    target.write(f"**{field}:** {item[field]}")

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

            c.drawString(50, h-60, "รายงานข้อมูลทรัพย์สิน")
            c.line(50, h-70, w-50, h-70)

            # QR Code ใน PDF
            qr_pdf_url = get_qr_url(data['ID-Auto'])
            if qr_pdf_url:
                qr_data = download_image(qr_pdf_url)
                if qr_data: c.drawImage(ImageReader(qr_data), w-130, h-130, 80, 80)

            # รายละเอียดข้อมูล
            c.setFont('ThaiBold', 14)
            y_pos = h - 100
            for f in FIELDS:
                if f not in ["รูปภาพ", "QR-CODE"]:
                    c.drawString(70, y_pos, f"• {f}: {data[f]}")
                    y_pos -= 22

            # รูปทรัพย์สินใน PDF
            img_pdf_url = get_drive_direct_link(data['รูปภาพ'])
            if img_pdf_url:
                img_data = download_image(img_pdf_url)
                if img_data: c.drawImage(ImageReader(img_data), 70, 50, width=250, height=180, preserveAspectRatio=True)

            c.save()
            return buf.getvalue()

        st.download_button(
            label="📥 ดาวน์โหลดรายงาน PDF",
            data=generate_pdf(item),
            file_name=f"Asset_Report_{item['ID-Auto']}.pdf",
            mime="application/pdf"
        )
