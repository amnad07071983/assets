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

# สำหรับสร้าง PDF
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

# --- 1. ตั้งค่า Google Sheets ---
def connect_sheet():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(creds)
        # เปลี่ยน 'Your_Sheet_Name' เป็นชื่อไฟล์ของคุณ
        return client.open("Your_Sheet_Name").worksheet("online")
    except Exception as e:
        messagebox.showerror("Error", f"เชื่อมต่อไม่สำเร็จ: {e}")
        return None

# รายชื่อฟิลด์ (17 ฟิลด์)
FIELDS = [
    "ID-Auto", "รูปภาพ", "QR-CODE", "บริษัท", "สถานะทรัพย์สิน", 
    "กลุ่มทรัพย์สิน", "รหัสทรัพย์สิน", "ชื่อทรัพย์สิน1", "แผนก", 
    "วันที่รับเข้าทะเบียน", "วันที่ตัดจากทะเบียน", "หน่วยนับ", 
    "จำนวน", "มูลค่าทุน", "ค่าเสื่อมสะสม", "มูลค่าคงเหลือ", "ข้อมูล ณ วันที่"
]

all_records = [] # เก็บข้อมูลทั้งหมดจาก Google Sheet
sheet = connect_sheet()

def load_data():
    global all_records
    if sheet:
        data = sheet.get_all_values()
        if data:
            # เก็บเฉพาะข้อมูล ไม่เอา Header แถวแรก
            all_records = [dict(zip(FIELDS, row)) for row in data[1:]]
            messagebox.showinfo("Success", f"โหลดข้อมูลสำเร็จ {len(all_records)} รายการ")
            apply_filter()

# --- 2. ฟังก์ชันกรองข้อมูล (Filter) ---
def parse_date(date_str):
    try:
        # พยายามแปลงวันที่ (รองรับ วว/ดด/ปปปป)
        return datetime.strptime(date_str.strip(), "%d/%m/%Y")
    except:
        return None

def apply_filter():
    search_id = entry_id.get().lower()
    search_comp = entry_company.get().lower()
    date_start = parse_date(entry_date_start.get())
    date_end = parse_date(entry_date_end.get())

    # ล้างข้อมูลเก่าในตาราง
    for item in tree.get_children():
        tree.delete(item)

    for row in all_records:
        # กรอง ID และ บริษัท (ค้นหาบางคำได้)
        match_id = search_id in row["ID-Auto"].lower()
        match_comp = search_comp in row["บริษัท"].lower()
        
        # กรองช่วงวันที่ (วันที่รับเข้าทะเบียน)
        row_date = parse_date(row["วันที่รับเข้าทะเบียน"])
        match_date = True
        if date_start and row_date:
            match_date = match_date and (row_date >= date_start)
        if date_end and row_date:
            match_date = match_date and (row_date <= date_end)

        if match_id and match_comp and match_date:
            tree.insert("", "end", values=(row["ID-Auto"], row["ชื่อทรัพย์สิน1"], row["บริษัท"], row["วันที่รับเข้าทะเบียน"]))

# --- 3. ฟังก์ชันสร้าง PDF ---
def export_selected_pdf():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Warning", "กรุณาเลือกรายการในตารางก่อนสั่งพิมพ์")
        return

    # ดึงข้อมูลจาก ID ที่เลือกในตาราง
    item_id = tree.item(selected_item)['values'][0]
    data = next((item for item in all_records if str(item["ID-Auto"]) == str(item_id)), None)

    if not data: return

    file_name = f"Report_{data['ID-Auto']}.pdf"
    c = canvas.Canvas(file_name, pagesize=A4)
    width, height = A4

    try:
        pdfmetrics.registerFont(TTFont('THSarabunBold', 'THSarabunNew Bold.ttf'))
        c.setFont('THSarabunBold', 18)
    except:
        messagebox.showerror("Error", "ไม่พบไฟล์ฟอนต์ THSarabunNew Bold.ttf")
        return

    # ส่วนบนขวา: QR-CODE
    try:
        qr_img = ImageReader(data['QR-CODE'])
        c.drawImage(qr_img, width - 130, height - 130, width=100, height=100)
    except:
        c.drawString(width - 130, height - 50, "[QR Error]")

    # รายละเอียดข้อมูล
    c.drawString(50, height - 50, f"รายงานทรัพย์สิน: {data['ชื่อทรัพย์สิน1']}")
    c.setFont('THSarabunBold', 14)
    y = height - 100
    
    for field in FIELDS:
        if field not in ["รูปภาพ", "QR-CODE"]:
            c.drawString(70, y, f"• {field}: {data.get(field, '-')}")
            y -= 22

    # ส่วนล่าง: รูปภาพทรัพย์สิน
    try:
        asset_img = ImageReader(data['รูปภาพ'])
        c.drawImage(asset_img, 70, y - 160, width=200, height=150, preserveAspectRatio=True)
    except:
        c.drawString(70, y - 50, "[ไม่สามารถโหลดรูปภาพได้]")

    c.save()
    messagebox.showinfo("Success", f"พิมพ์ไฟล์ {file_name} เรียบร้อย!")

# --- 4. UI Setup ---
root = tk.Tk()
root.title("System V2 (Filter & PDF)")
root.geometry("800x750")

# Filter Section
f_frame = tk.LabelFrame(root, text="ค้นหาและกรองข้อมูล", padx=10, pady=10)
f_frame.pack(fill="x", padx=10, pady=5)

tk.Label(f_frame, text="ค้นหา ID:").grid(row=0, column=0)
entry_id = tk.Entry(f_frame)
entry_id.grid(row=0, column=1, padx=5)

tk.Label(f_frame, text="ค้นหาบริษัท:").grid(row=0, column=2)
entry_company = tk.Entry(f_frame)
entry_company.grid(row=0, column=3, padx=5)

tk.Label(f_frame, text="วันที่เริ่ม (วว/ดด/ปปปป):").grid(row=1, column=0, pady=5)
entry_date_start = tk.Entry(f_frame)
entry_date_start.grid(row=1, column=1)

tk.Label(f_frame, text="วันที่สิ้นสุด (วว/ดด/ปปปป):").grid(row=1, column=2)
entry_date_end = tk.Entry(f_frame)
entry_date_end.grid(row=1, column=3)

btn_filter = tk.Button(f_frame, text="🔍 กรองข้อมูล", command=apply_filter, bg="#eee")
btn_filter.grid(row=1, column=4, padx=10)

# Table Section
tree_frame = tk.Frame(root)
tree_frame.pack(fill="both", expand=True, padx=10)

columns = ("ID", "Name", "Company", "DateIn")
tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
tree.heading("ID", text="ID-Auto")
tree.heading("Name", text="ชื่อทรัพย์สิน")
tree.heading("Company", text="บริษัท")
tree.heading("DateIn", text="วันที่รับเข้า")
tree.column("ID", width=80)
tree.pack(side="left", fill="both", expand=True)

# Action Buttons
action_frame = tk.Frame(root)
action_frame.pack(pady=10)

btn_reload = tk.Button(action_frame, text="🔄 โหลดข้อมูลจาก Google", command=load_data)
btn_reload.pack(side="left", padx=5)

btn_print = tk.Button(action_frame, text="🖨️ พิมพ์ PDF รายการที่เลือก", bg="#2196F3", fg="white", command=export_selected_pdf)
btn_print.pack(side="left", padx=5)

# โหลดข้อมูลทันทีเมื่อเปิดโปรแกรม
load_data()

root.mainloop()
