import cv2
import gspread
import requests
from oauth2client.service_account import ServiceAccountCredentials
from pyzbar.pyzbar import decode
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import messagebox
from io import BytesIO

# --- 1. การตั้งค่า Google Sheets ---
def connect_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    # เปลี่ยน 'credentials.json' เป็นชื่อไฟล์ของคุณ
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)
    # เปลี่ยน 'Your_Sheet_Name' เป็นชื่อไฟล์ Google Sheet ของคุณ
    return client.open("Your_Sheet_Name").sheet1

sheet = connect_sheet()

# --- 2. ฟังก์ชันดึงข้อมูลและแสดงผลบน UI ---
def fetch_and_display(search_id):
    try:
        # ค้นหา ID ในคอลัมน์ A (คอลัมน์ที่ 1)
        cell = sheet.find(search_id)
        data = sheet.row_values(cell.row)
        
        # สมมติลำดับคือ: [ID, Name, Image_URL]
        res_id, res_name, res_url = data[0], data[1], data[2]
        
        # อัปเดตข้อความบนหน้าจอ
        label_info.config(text=f"ID: {res_id}\nName: {res_name}", fg="black")
        
        # ดึงรูปภาพจาก URL
        response = requests.get(res_url)
        img_data = Image.open(BytesIO(response.content))
        img_data = img_data.resize((250, 250)) # ปรับขนาดรูป
        img_tk = ImageTk.PhotoImage(img_data)
        
        label_img.config(image=img_tk)
        label_img.image = img_tk # ป้องกัน Garbage Collection

    except Exception as e:
        label_info.config(text="❌ ไม่พบข้อมูล หรือเกิดข้อผิดพลาด", fg="red")
        print(f"Error: {e}")

# --- 3. ฟังก์ชันเปิดกล้องแสกน QR Code ---
def start_scan():
    cap = cv2.VideoCapture(0)
    while True:
        ret, frame = cap.read()
        for barcode in decode(frame):
            code_data = barcode.data.decode('utf-8')
            entry_id.delete(0, tk.END)
            entry_id.insert(0, code_data)
            fetch_and_display(code_data) # ส่งไปค้นหาทันที
            cap.release()
            cv2.destroyAllWindows()
            return

        cv2.imshow("QR Scanner (Press 'q' to exit)", frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    cap.release()
    cv2.destroyAllWindows()

# --- 4. สร้างหน้าจอ GUI ด้วย Tkinter ---
root = tk.Tk()
root.title("Google Sheets Scanner System")
root.geometry("400x550")

# ส่วนคีย์มือ
tk.Label(root, text="กรอก ID หรือ แสกน QR Code", font=("Arial", 12)).pack(pady=10)
entry_id = tk.Entry(root, font=("Arial", 14), justify='center')
entry_id.pack(pady=5)

btn_search = tk.Button(root, text="ค้นหาด้วย ID", command=lambda: fetch_and_display(entry_id.get()))
btn_search.pack(pady=5)

btn_scan = tk.Button(root, text="📷 เปิดกล้องแสกน", bg="#4CAF50", fg="white", command=start_scan)
btn_scan.pack(pady=5)

# ส่วนแสดงผล
label_info = tk.Label(root, text="รอรับข้อมูล...", font=("Arial", 14, "bold"), pady=20)
label_info.pack()

label_img = tk.Label(root) # พื้นที่โชว์รูป
label_img.pack(pady=10)

root.mainloop()
