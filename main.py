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

# --- 1. ฟังก์ชันช่วยดึง URL ออกจากสูตร =IMAGE และแปลงเป็น Direct Link ---
def get_drive_direct_link(cell_value):
    if not cell_value:
        return None
    
    # 1.1 ถ้ามีสูตร =IMAGE("...") ให้ดึงเฉพาะ URL ออกมา
    url_match = re.search(r'https?://[^\s"]+', cell_value)
    if not url_match:
        return None
    
    url = url_match.group(0)
    
    # 1.2 ถ้าเป็นลิงก์ Google Drive ให้แปลงเป็น Direct Link สำหรับ Download
    if "drive.google.com" in url:
        # ดึง File ID ออกจากลิงก์
        file_id = ""
        if "/d/" in url:
            file_id = url.split("/d/")[1].split("/")[0].split("?")[0]
        elif "id=" in url:
            file_id = url.split("id=")[1].split("&")[0]
        
        if file_id:
            return f"https://drive.google.com/uc?export=download&id={file_id}"
    
    return url

# --- 2. ฟังก์ชันดาวน์โหลดรูปภาพ ---
def download_image(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    try:
        response = requests.get(url, headers=headers, timeout=15)
        if response.status_code == 200:
            return BytesIO(response.content)
    except Exception as e:
        st.error(f"Error downloading image: {e}")
    return None

# --- ส่วนของการสร้าง PDF (เฉพาะในฟังก์ชัน generate_pdf) ---
def generate_pdf(data):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    
    try:
        pdfmetrics.registerFont(TTFont('ThaiBold', 'THSARABUN BOLD.ttf'))
        c.setFont('ThaiBold', 22)
    except:
        c.setFont('Helvetica-Bold', 18)

    # วาดหัวข้อและข้อมูลตัวอักษร... (ส่วนเดิมของคุณ)
    c.drawString(50, h-60, "รายงานข้อมูลทรัพย์สิน")
    c.line(50, h-70, w-150, h-70)

    # จัดการ QR-CODE
    qr_link = get_drive_direct_link(data['QR-CODE'])
    if qr_link:
        qr_data = download_image(qr_link)
        if qr_data:
            try:
                c.drawImage(ImageReader(qr_data), w-130, h-130, 100, 100)
            except: pass

    # วาดฟิลด์ข้อมูลอื่นๆ...
    c.setFont('ThaiBold', 15)
    curr_y = h - 110
    for f in FIELDS:
        if f not in ["รูปภาพ", "QR-CODE"]:
            c.drawString(70, curr_y, f"• {f}: {data[f]}")
            curr_y -= 24

    # จัดการ รูปทรัพย์สิน (ด้านล่าง)
    img_link = get_drive_direct_link(data['รูปภาพ'])
    if img_link:
        img_data = download_image(img_link)
        if img_data:
            try:
                # ใช้ ImageReader ครอบ BytesIO
                c.drawImage(ImageReader(img_data), 70, 50, width=280, height=200, preserveAspectRatio=True)
            except Exception as e:
                c.setFont('ThaiBold', 10)
                c.drawString(70, 50, f"(ไม่สามารถแสดงรูปภาพได้ใน PDF: {e})")

    c.save()
    return buf.getvalue()

# (โค้ดส่วนที่เหลือใน main.py ให้ใช้ตามเดิม แต่เปลี่ยนฟังก์ชัน generate_pdf เป็นตัวใหม่นี้ครับ)
