import cv2
import gspread
import requests
import tkinter as tk
from tkinter import messagebox, ttk
from oauth2client.service_account import ServiceAccountCredentials
from pyzbar.pyzbar import decode
from PIL import Image, ImageTk
from io import BytesIO

# สำหรับสร้าง PDF
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

# --- 1. การตั้งค่า Google Sheets ---
def connect_sheet():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(creds)
        # ระบุชื่อไฟล์และชื่อแผ่นงาน "online"
        return client.open("Your_Sheet_Name").worksheet("online")
    except Exception as e:
        messagebox.showerror("Error", f"เชื่อมต่อ Sheet ไม่ได้: {e}")
        return None

sheet = connect_sheet()

# ฟิลด์ข้อมูลทั้งหมด (17 ฟิลด์)
FIELDS = [
    "ID-Auto", "รูปภาพ", "QR-CODE", "บริษัท", "สถานะทรัพย์สิน", 
    "กลุ่มทรัพย์สิน", "รหัสทรัพย์สิน", "ชื่อทรัพย์สิน1", "แผนก", 
    "วันที่รับเข้าทะเบียน", "วันที่ตัดจากทะเบียน", "หน่วยนับ", 
    "จำนวน", "มูลค่าทุน", "ค่าเสื่อมสะสม", "มูลค่าคงเหลือ", "ข้อมูล ณ วันที่"
]

current_data = None # เก็บข้อมูลล่าสุดที่ค้นหาเจอ

# --- 2. ฟังก์ชันดึงข้อมูลและแสดงผล ---
def fetch_and_display(search_id):
    global current_data
    try:
        # ค้นหา ID ในคอลัมน์ A
        cell = sheet.find(str(search_id))
        row_data = sheet.row_values(cell.row)
        
        # ตรวจสอบว่าข้อมูลมีครบตามจำนวนฟิลด์หรือไม่ (เผื่อกรณีแถวไม่สมบูรณ์)
        while len(row_data) < len(FIELDS):
            row_data.append("-")
            
        current_data = dict(zip(FIELDS, row_data))
        
        # อัปเดต UI (แสดงข้อมูลหลักๆ)
        info_text = (f"ID: {current_data['ID-Auto']}\n"
                     f"ชื่อ: {current_data['ชื่อทรัพย์สิน1']}\n"
                     f"บริษัท: {current_data['บริษัท']}\n"
                     f"รหัสทรัพย์สิน: {current_data['รหัสทรัพย์สิน']}")
        label_info.config(text=info_text, fg="black")

        # แสดงรูปภาพ (ลำดับที่ 1 ใน list คือ รูปภาพ)
        try:
            res_url = current_data['รูปภาพ']
            response = requests.get(res_url, timeout=5)
            img_data = Image.open(BytesIO(response.content))
            img_data = img_data.resize((150, 150))
            img_tk = ImageTk.PhotoImage(img_data)
            label_img.config(image=img_tk)
            label_img.image = img_tk
        except:
            label_img.config(image='', text="[ไม่มีรูปภาพ]")

    except Exception as e:
        label_info.config(text="❌ ไม่พบข้อมูล", fg="red")
        current_data = None

# --- 3. ฟังก์ชันสร้าง PDF (ตามเงื่อนไขที่กำหนด) ---
def export_pdf():
    if not current_data:
        messagebox.showwarning("Warning", "กรุณาค้นหาข้อมูลก่อนสั่งพิมพ์")
        return

    file_name = f"Report_{current_data['ID-Auto']}.pdf"
    c = canvas.Canvas(file_name, pagesize=A4)
    width, height = A4

    # ลงทะเบียนฟอนต์ไทย
    try:
        pdfmetrics.registerFont(TTFont('THSarabunBold', 'THSarabunNew Bold.ttf'))
        c.setFont('THSarabunBold', 16)
    except:
        messagebox.showerror("Error", "ไม่พบไฟล์ฟอนต์ THSarabunNew Bold.ttf")
        return

    # 1. แสดง QR-CODE ด้านบนขวามือ
    try:
        qr_url = current_data['QR-CODE']
        qr_img = ImageReader(qr_url)
        c.drawImage(qr_img, width - 120, height - 120, width=100, height=100)
    except:
        c.drawString(width - 120, height - 50, "[QR Error]")

    # 2. รายละเอียดข้อมูล (เรียงลำดับลงมา)
    c.drawString(50, height - 50, f"รายงานข้อมูลทรัพย์สิน")
    y_position = height - 100
    
    for field in FIELDS:
        if field not in ["รูปภาพ", "QR-CODE"]:
            text = f"{field}: {current_data.get(field, '-')}"
            c.drawString(50, y_position, text)
            y_position -= 25

    # 3. แสดงรูปภาพที่ด้านล่างของรายงาน
    try:
        main_img_url = current_data['รูปภาพ']
        main_img = ImageReader(main_img_url)
        c.drawImage(main_img, 50, y_position - 150, width=200, height=150)
    except:
        c.drawString(50, y_position - 50, "[ไม่สามารถโหลดรูปภาพประกอบได้]")

    c.save()
    messagebox.showinfo("Success", f"สร้างไฟล์ {file_name} เรียบร้อยแล้ว")

# --- 4. ฟังก์ชันเปิดกล้องแสกน ---
def start_scan():
    cap = cv2.VideoCapture(0)
    while True:
        ret, frame = cap.read()
        if not ret: break
        for barcode in decode(frame):
            code_data = barcode.data.decode('utf-8')
            entry_id.delete(0, tk.END)
            entry_id.insert(0, code_data)
            fetch_and_display(code_data)
            cap.release()
            cv2.destroyAllWindows()
            return
        cv2.imshow("QR Scanner (Press 'q' to exit)", frame)
        if cv2.waitKey(1) & 0xFF == ord('q'): break
    cap.release()
    cv2.destroyAllWindows()

# --- 5. สร้าง GUI ---
root = tk.Tk()
root.title("Asset Management System")
root.geometry("4500x700")

# ส่วน Filter (กรองข้อมูล)
filter_frame = tk.LabelFrame(root, text="ตัวกรองข้อมูล (Filter)", padx=10, pady=10)
filter_frame.pack(pady=10, fill="x", padx=20)

tk.Label(filter_frame, text="ID-Auto:").grid(row=0, column=0)
entry_id = tk.Entry(filter_frame)
entry_id.grid(row=0, column=1)

tk.Label(filter_frame, text="บริษัท:").grid(row=0, column=2)
entry_company = tk.Entry(filter_frame)
entry_company.grid(row=0, column=3)

# ปุ่มควบคุม
btn_frame = tk.Frame(root)
btn_frame.pack(pady=10)

btn_search = tk.Button(btn_frame, text="🔍 ค้นหา/กรอง", command=lambda: fetch_and_display(entry_id.get()))
btn_search.pack(side="left", padx=5)

btn_scan = tk.Button(btn_frame, text="📷 แสกน QR", bg="#4CAF50", fg="white", command=start_scan)
btn_scan.pack(side="left", padx=5)

btn_print = tk.Button(btn_frame, text="🖨️ พิมพ์ PDF", bg="#2196F3", fg="white", command=export_pdf)
btn_print.pack(side="left", padx=5)

# ส่วนแสดงผลบนจอ
label_info = tk.Label(root, text="รอรับข้อมูล...", font=("Arial", 12), justify="left")
label_info.pack(pady=10)

label_img = tk.Label(root)
label_img.pack(pady=10)

root.mainloop()
