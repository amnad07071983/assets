import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from PIL import Image
from io import BytesIO
from datetime import datetime
import requests  # เพิ่มการนำเข้า requests
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

# --- 2. ฟังก์ชันเชื่อมต่อ Google Sheets ---
@st.cache_resource
def get_data_from_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_info = st.secrets["gcp_service_account"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
        client = gspread.authorize(creds)
        
        # ID ของ Google Sheet
        SHEET_ID = "1Pp2XffqRBtlyDu6NHDmFA6VcbdCmZPEv-1p3ETCSb5o" 
        
        sh = client.open_by_key(SHEET_ID) 
        worksheet = sh.worksheet("online")
        
        data = worksheet.get_all_values()
        if len(data) > 1:
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
    search_id = st.sidebar.text_input("ค้นหา ID-Auto", placeholder="พิมพ์ ID...")
    search_comp = st.sidebar.text_input("ค้นหา บริษัท", placeholder="พิมพ์ชื่อบริษัท...")
    
    st.sidebar.subheader("📅 กรองด้วยวันที่")
    date_in = st.sidebar.text_input("วันที่รับเข้า (วว/ดด/ปปปป)", placeholder="เช่น 01/01/2024")
    date_out = st.sidebar.text_input("วันที่ตัดทะเบียน (วว/ดด/ปปปป)", placeholder="เช่น 31/12/2024")

    filtered_df = df.copy()
    if search_id:
        filtered_df = filtered_df[filtered_df['ID-Auto'].str.contains(search_id, case=False, na=False)]
    if search_comp:
        filtered_df = filtered_df[filtered_df['บริษัท'].str.contains(search_comp, case=False, na=False)]
    if date_in:
        filtered_df = filtered_df[filtered_df['วันที่รับเข้าทะเบียน'].str.contains(date_in, na=False)]
    if date_out:
        filtered_df = filtered_df[filtered_df['วันที่ตัดจากทะเบียน'].str.contains(date_out, na=False)]

    st.write(f"📊 พบข้อมูลทั้งหมด **{len(filtered_df)}** รายการ")
    
    selection = st.dataframe(
        filtered_df, 
        use_container_width=True, 
        on_select="rerun", 
        selection_mode="single-row",
        hide_index=True
    )

    # --- 6. แสดงรายละเอียดและการดาวน์โหลด PDF ---
    if len(selection.selection.rows) > 0:
        row_idx = selection.selection.rows[0]
        item = filtered_df.iloc[row_idx]
        
        st.divider()
        col_img, col_detail = st.columns([1, 2])
        
        with col_img:
            if item['รูปภาพ']:
                st.image(item['รูปภาพ'], caption="📸 รูปทรัพย์สิน", use_container_width=True)
            if item['QR-CODE']:
                st.image(item['QR-CODE'], caption="🔗 QR-CODE", width=150)
            
        with col_detail:
            st.subheader(f"📄 รายละเอียดทรัพย์สิน: {item['ชื่อทรัพย์สิน1']}")
            info_1, info_2 = st.columns(2)
            for i, field in enumerate(FIELDS):
                if field not in ["รูปภาพ", "QR-CODE"]:
                    target = info_1 if i % 2 == 0 else info_2
                    target.write(f"**{field}:** {item[field]}")

            # --- 7. ฟังก์ชันสร้าง PDF (แก้ไขส่วนดึงรูปภาพ) ---
            def generate_pdf(data):
                buf = BytesIO()
                c = canvas.Canvas(buf, pagesize=A4)
                w, h = A4
                
                try:
                    pdfmetrics.registerFont(TTFont('ThaiBold', 'THSARABUN BOLD.ttf'))
                    c.setFont('ThaiBold', 22)
                except:
                    c.setFont('Helvetica-Bold', 18)

                # 1. วาด QR-CODE (ดึงผ่าน requests)
                if data['QR-CODE']:
                    try:
                        resp_qr = requests.get(data['QR-CODE'], timeout=10)
                        qr_img = ImageReader(BytesIO(resp_qr.content))
                        c.drawImage(qr_img, w-130, h-130, 100, 100)
                    except Exception as e:
                        pass # ถ้าโหลดไม่ได้ให้ข้ามไป

                c.drawString(50, h-60, "รายงานข้อมูลทรัพย์สิน")
                c.setLineWidth(1)
                c.line(50, h-70, w-150, h-70)
                
                c.setFont('ThaiBold', 15)
                curr_y = h - 110
                for f in FIELDS:
                    if f not in ["รูปภาพ", "QR-CODE"]:
                        c.drawString(70, curr_y, f"• {f}: {data[f]}")
                        curr_y -= 24
                
                # 2. วาดรูปทรัพย์สิน (ดึงผ่าน requests)
                if data['รูปภาพ']:
                    try:
                        resp_asset = requests.get(data['รูปภาพ'], timeout=10)
                        asset_img = ImageReader(BytesIO(resp_asset.content))
                        # วาดรูปไว้ด้านล่าง
                        c.drawImage(asset_img, 70, 50, width=280, height=200, preserveAspectRatio=True)
                    except Exception as e:
                        c.setFont('ThaiBold', 10)
                        c.drawString(70, 50, f"(ไม่สามารถโหลดรูปภาพใน PDF ได้)")
                
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
    st.info("💡 กรุณาตรวจสอบการตั้งค่า Sheet ID และ Secrets ให้เรียบร้อย")
