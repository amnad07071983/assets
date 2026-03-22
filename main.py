import cv2
import gspread
import requests
import tkinter as tk
from tkinter import messagebox, ttk
from oauth2client.service_account import ServiceAccountCredentials
from pyzbar.pyzbar import decode
from PIL import Image, ImageTk
from io import BytesIO
from datetime import datetime
import os

# สำหรับสร้าง PDF
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

# --- 1. ตั้งค่าการเชื่อมต่อ Google Sheets ---
def connect_sheet():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        # ต้องมีไฟล์ credentials.json ใน GitHub ของคุณด้วย
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(creds)
        # เปลี่ยนชื่อไฟล์ให้ตรงกับของคุณ
        return client.open("Your_Sheet_Name").worksheet("online")
    except Exception as e:
        print(f"Error connecting: {e}")
        return None

FIELDS = [
    "ID-Auto", "รูปภาพ", "QR-CODE", "บริษัท", "สถานะทรัพย์สิน", 
    "กลุ่มทรัพย์สิน", "รหัสทรัพย์สิน", "ชื่อทรัพย์สิน1", "แผนก", 
    "วันที่รับเข้าทะเบียน", "วันที่ตัดจากทะเบียน", "หน่วยนับ", 
    "จำนวน", "มูลค่าทุน", "ค่าเสื่อมสะสม", "มูลค่าคงเหลือ", "ข้อมูล ณ วันที่"
]

all_records = []
sheet = connect_sheet()

# --- 2. ฟังก์ชันจัดการข้อมูลและ Filter ---
def load_data():
    global all_records
    if sheet:
        data = sheet.get_all_values()
        if data:
            all_records = [dict(zip(FIELDS, row)) for row in data[1:]]
            apply_filter()

def parse_date(date_str):
    for fmt in ("%d/%m/%Y", "%d/%m/%y"): # รองรับทั้ง 2024 และ 24
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except:
            continue
    return None

def apply_filter():
    search_id = entry_id.get().lower()
    search_comp = entry_company.get().lower()
    date_start = parse_date(entry_date_start.get())
    date_end = parse_date(entry_date_end.get())

    for item in tree.get_children():
        tree.delete(item)

    for row in all_records:
        # Filter ข้อความ (Search บางคำได้)
        match_id = search_id in str(row["ID-Auto"]).lower()
        match_comp = search_comp in str(row["บริษัท"]).lower()
        
        # Filter วันที่
        row_date = parse_date(row["วันที่รับเข้าทะเบียน"])
        match_date = True
        if date_start and row_date: match_date = match_date and (row_date >= date_start)
        if date_end and row_date: match_date = match_date and (row_date <= date_end)

        if match_id and match_comp and match_date:
            tree.insert("", "end", values=(row["ID-Auto"], row["ชื่อทรัพย์สิน1"], row["บริษัท"], row["วันที่รับเข้าทะเบียน"]))

# --- 3. ฟังก์ชันสร้าง PDF และ Download ---
def export_pdf():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("คำเตือน", "กรุณาเลือกรายการที่จะพิมพ์")
        return

    item_id = tree.item(selected)['values'][0]
    data = next((item for item in all_records if str(item["ID-Auto"]) == str(item_id)), None)
    
    if not data: return

    file_name = f"Asset_Report_{data['ID-Auto']}.pdf"
    c = canvas.Canvas(file_name, pagesize=A4)
    width, height = A4

    # ใช้ฟอนต์จากไฟล์ที่คุณมีใน GitHub
    font_path = "THSARABUN BOLD.ttf"
    if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont('ThaiFontBold', font_path))
        c.setFont('ThaiFontBold', 16)
    else:
        c.setFont('Helvetica-Bold', 16)

    # 1. QR-CODE มุมบนขวา
    try:
        qr_url = data['QR-CODE']
        qr_img = ImageReader(qr_url)
        c.drawImage(qr_img, width - 130, height - 130, width=100, height=100)
    except:
        pass

    # 2. รายละเอียดข้อมูล
    c.drawString(50, height - 50, f"รายงานข้อมูลทรัพย์สิน: {data['ชื่อทรัพย์สิน1']}")
    c.setLineWidth(1)
    c.line(50, height - 60, width - 50, height - 60)
    
    y = height - 100
    for field in FIELDS:
        if field not in ["รูปภาพ", "QR-CODE"]:
            c.drawString(70, y, f"• {field}: {data.get(field, '-')}")
            y -= 25

    # 3. รูปภาพด้านล่างของรายงาน
    try:
        asset_url = data['รูปภาพ']
        asset_img = ImageReader(asset_url)
        # ปรับขนาดรูปให้พอดี
        c.drawImage(asset_img, 70, 50, width=250, height=180, preserveAspectRatio=True)
    except:
        c.drawString(70, 100, "[ไม่สามารถโหลดรูปภาพทรัพย์สินได้]")

    c.save()
    messagebox.showinfo("สำเร็จ", f"สร้างไฟล์ {file_name} แล้วในโฟลเดอร์แอป")

# --- 4. ส่วนประกอบของหน้าจอ (UI) ---
root = tk.Tk()
root.title("ระบบจัดการทรัพย์สิน (GitHub Version)")
root.geometry("800x650")

# ส่วนกรองข้อมูล
filter_frame = tk.LabelFrame(root, text=" 🔍 ตัวกรองข้อมูล ", padx=10, pady=10)
filter_frame.pack(fill="x", padx=15, pady=10)

tk.Label(filter_frame, text="ID / บางคำ:").grid(row=0, column=0)
entry_id = tk.Entry(filter_frame)
entry_id.grid(row=0, column=1, padx=5)

tk.Label(filter_frame, text="บริษัท:").grid(row=0, column=2)
entry_company = tk.Entry(filter_frame)
entry_company.grid(row=0, column=3, padx=5)

tk.Label(filter_frame, text="วันที่เริ่ม (วว/ดด/ปปปป):").grid(row=1, column=0, pady=5)
entry_date_start = tk.Entry(filter_frame)
entry_date_start.grid(row=1, column=1)

tk.Label(filter_frame, text="ถึงวันที่:").grid(row=1, column=2)
entry_date_end = tk.Entry(filter_frame)
entry_date_end.grid(row=1, column=3)

tk.Button(filter_frame, text="กรองข้อมูล", command=apply_filter, bg="#eee").grid(row=1, column=4, padx=10)

# ส่วนตารางแสดงผล
tree = ttk.Treeview(root, columns=("ID", "Name", "Comp", "Date"), show="headings")
tree.heading("ID", text="ID-Auto"); tree.heading("Name", text="ชื่อทรัพย์สิน")
tree.heading("Comp", text="บริษัท"); tree.heading("Date", text="วันที่รับเข้า")
tree.pack(fill="both", expand=True, padx=15)

# ส่วนปุ่มคำสั่ง
btn_frame = tk.Frame(root)
btn_frame.pack(pady=15)

tk.Button(btn_frame, text="🔄 โหลดข้อมูลใหม่", command=load_data, padx=10).pack(side="left", padx=5)
tk.Button(btn_frame, text="🖨️ ดาวน์โหลด PDF รายการที่เลือก", bg="#2196F3", fg="white", command=export_pdf, padx=10).pack(side="left", padx=5)

load_data()
root.mainloop()
