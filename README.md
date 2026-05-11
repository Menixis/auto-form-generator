---
title: Auto Form Generator
emoji: 📄
colorFrom: blue
colorTo: indigo
sdk: docker
app_port: 7860
pinned: false
short_description: ระบบสร้างหนังสือราชการอัตโนมัติ - ปตท.
---
# Auto Form Generator - ปตท.

ระบบสร้างหนังสือราชการอัตโนมัติจากแบบฟอร์มมาตรฐาน ปตท. 5 ประเภท พร้อมดึงข้อมูล PO อัตโนมัติจากไฟล์ Excel

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

---

## ✨ Features

- 🖥️ **GUI ภาษาไทย** ใช้งานง่ายด้วย tkinter
- 📄 **รองรับ 5 แบบฟอร์ม** มาตรฐาน ปตท.
- 🔍 **Auto-fill จาก Excel** — ใส่ PO แล้วระบบเติม Vendor, ชื่องาน ฯลฯ ให้อัตโนมัติ
- 📂 **รองรับ OneDrive/SharePoint** — เปิดไฟล์ Excel ที่ sync จาก cloud ได้ตรงๆ
- ✏️ **รักษารูปแบบเอกสาร** — Font, สี, layout ของต้นฉบับ ครบทุกอย่าง
- 🚀 **ไม่ต้องใช้ API หรือสิทธิ์ admin** — รันได้บนเครื่องส่วนตัวเลย

---

## 📋 แบบฟอร์มที่รองรับ

| # | ชื่อแบบฟอร์ม |
|---|---|
| 1 | หนังสือแจ้งให้เริ่มเข้าทำงาน / ส่งมอบผลิตภัณฑ์ (ครั้งแรก) |
| 2 | หนังสือแจ้งให้กลับเข้าทำงาน (หลังแจ้งหยุดงาน) |
| 3 | หนังสือแจ้งหยุดงาน (มิใช่ความผิดของ ปตท. และคู่สัญญา) |
| 4 | หนังสือแจ้งหยุดงาน (เหตุจากความผิดของคู่สัญญา) |
| 5 | หนังสือแจ้งหยุดงาน (เหตุจากความผิดของ ปตท.) |

---

## 📦 Requirements

- **Python 3.8** ขึ้นไป ([Download](https://www.python.org/downloads/))
- **Microsoft Office** หรือ **LibreOffice** สำหรับเปิดเอกสารที่สร้าง
- **OneDrive / SharePoint sync** (ถ้าใช้ไฟล์ Excel ที่อยู่บน cloud)

Python packages:
- python-docx >= 1.0.0
- openpyxl >= 3.1.0

---

## 🚀 Quick Start

### 1. Clone repository

```bash
git clone https://github.com/<your-username>/auto-form-generator.git
cd auto-form-generator
```

### 2. ติดตั้ง dependencies

```bash
pip install -r requirements.txt
```

หรือใช้ virtual environment (แนะนำ):

```bash
python -m venv venv
# Windows
venv\Scripts\activate
# macOS / Linux
source venv/bin/activate

pip install -r requirements.txt
```

### 3. ตั้งค่าไฟล์ Excel

คัดลอก template:
```bash
# Windows
copy config.ini.template config.ini
# macOS / Linux
cp config.ini.template config.ini
```

แก้ไข `config.ini` ให้ชี้ไปที่ไฟล์ Excel ของคุณ:

```ini
[excel]
file_path = C:\Users\YourName\OneDrive - ปตท\PO_Data\po_data.xlsx
sheet_name = 
po_column = poNo
```

### 4. รันโปรแกรม

```bash
python auto_form_generator.py
```

---

## ⚙️ Configuration

### `config.ini` — Excel data source

| Key | คำอธิบาย | ตัวอย่าง |
|---|---|---|
| `file_path` | path ของไฟล์ Excel (.xlsx) | `C:\Users\xxx\OneDrive\po_data.xlsx` |
| `sheet_name` | ชื่อ sheet (เว้นว่างเพื่อใช้ sheet แรก) | `Sheet1` หรือเว้นว่าง |
| `po_column` | ชื่อคอลัมน์ที่เก็บเลข PO | `poNo` |

> ⚠️ **อย่า commit `config.ini` เข้า git** — มี `.gitignore` ป้องกันให้แล้ว

### โครงสร้าง Excel ที่รองรับ

แถวแรกเป็น header แถวต่อๆไปเป็นข้อมูล มีคอลัมน์อย่างน้อย:

| poNo | VendorName | title | prDocument | ... |
|---|---|---|---|---|
| 4500987654 | บริษัท ABC จำกัด | งานปรับปรุงระบบ | PR-2568-0001 | ... |

โปรแกรมจะ auto-fill ข้อมูลเหล่านี้ลงในแบบฟอร์ม:
- `VendorName` → "Vendor" / "[ระบุชื่อคู่สัญญา]"
- `title` → "ชื่องาน"

---

## 📖 วิธีใช้งาน

1. **เลือกประเภทแบบฟอร์ม** จาก dropdown
2. **ใส่เลข PO** → กดปุ่ม **"ดึงข้อมูลจาก Excel"** (สีเขียว)
   - ข้อมูลที่ดึงมาจะแสดงในกล่อง preview
3. **กรอกข้อมูลที่เหลือ**:
   - หมายเลขเอกสาร (เช่น `PTT-001/2568`)
   - ประเภทเอกสาร (เริ่มเข้าทำงาน / เริ่มส่งมอบสินค้า) — เฉพาะแบบฟอร์ม 1
   - ผู้ลงนาม
4. **กด "สร้างเอกสาร"** — ระบบจะสร้างไฟล์ใน `output/` พร้อม timestamp

### ทดสอบจาก command line ก่อนเปิด GUI

```bash
# ดูโครงสร้างไฟล์ Excel
python excel_lookup.py --list

# ทดสอบค้นหา PO
python excel_lookup.py 4500987654
```

---

## 📁 Project Structure

```
auto-form-generator/
├── .gitignore                    # ไฟล์ที่ไม่ commit (config, output, cache)
├── LICENSE                       # MIT License
├── README.md                     # ไฟล์นี้
├── requirements.txt              # Python dependencies
├── config.ini.template           # ตัวอย่าง config (คัดลอกเป็น config.ini)
├── auto_form_generator.py        # โปรแกรมหลัก (GUI)
├── excel_lookup.py               # โมดูลดึงข้อมูลจาก Excel
├── sample_po_data.xlsx           # ไฟล์ตัวอย่างสำหรับทดสอบ
├── templates/                    # แบบฟอร์มต้นฉบับ
│   ├── form1.docx
│   ├── form2.docx
│   ├── form3.docx
│   ├── form4.docx
│   └── form5.docx
└── output/                       # ไฟล์ที่สร้าง (auto-created, gitignored)
```

---

## 🔧 Customization

### ถ้า Excel ของคุณใช้ชื่อคอลัมน์อื่น

**กรณีที่ 1: คอลัมน์ PO ใช้ชื่ออื่น** (เช่น `PO Number` แทน `poNo`)

แก้ใน `config.ini`:
```ini
po_column = PO Number
```

**กรณีที่ 2: คอลัมน์ Vendor หรือ Title ใช้ชื่ออื่น**

แก้ใน `auto_form_generator.py` ที่ฟังก์ชัน `generate_form()`:

```python
vendor_name = (lookup_data or {}).get("Vendor_Name") or ""   # เปลี่ยน key
work_title = (lookup_data or {}).get("Job_Title") or ""      # เปลี่ยน key
```

และที่ตัวแปร `DISPLAY_FIELDS`:
```python
DISPLAY_FIELDS = [
    ("poNo", "เลข PO"),
    ("Vendor_Name", "ชื่อคู่สัญญา"),      # เปลี่ยน key ตรงนี้
    ("Job_Title", "ชื่องาน"),            # และตรงนี้
    ...
]
```

### เพิ่ม placeholder ใหม่ที่ auto-fill จาก Excel

ใน `auto_form_generator.py` ฟังก์ชัน `generate_form()`:

```python
# เช่น เพิ่ม auto-fill จำนวนวัน
day_sla = (lookup_data or {}).get("daySLA")
if day_sla:
    n = replace_text(doc, "จำนวนวัน", str(day_sla))
    log.append(f"จำนวนวัน (Excel): แทนที่ {n} จุด")
```

---

## 🐛 Troubleshooting

| ปัญหา | สาเหตุ | วิธีแก้ |
|---|---|---|
| `ไม่พบไฟล์ Excel` | path ผิด / ไฟล์ยังไม่ sync | ตรวจ path ใน config.ini, คลิกขวาที่ไฟล์ใน OneDrive → "Always keep on this device" |
| `ไม่สามารถเปิดไฟล์ได้` | ไฟล์ถูกเปิดอยู่ในโปรแกรมอื่น | ปิด Excel ก่อนรัน |
| `ไม่พบคอลัมน์ 'poNo'` | header ใน Excel ไม่ตรงกับ config | แก้ `po_column` ใน config.ini หรือเปลี่ยน header ใน Excel |
| `ไม่พบ PO: xxx` | PO ไม่มีในไฟล์ Excel | ตรวจดูว่า PO มีจริงในไฟล์ (run `python excel_lookup.py --list`) |
| `ImportError: No module named 'openpyxl'` | ยังไม่ได้ติดตั้ง dependencies | `pip install -r requirements.txt` |
| `ImportError: No module named 'tkinter'` | Python ไม่ได้มาพร้อม tkinter (rare) | Windows: ติดตั้ง Python ใหม่จาก python.org / Linux: `sudo apt install python3-tk` |
| GUI font ภาษาไทยเพี้ยน | OS ไม่มี font Tahoma | macOS/Linux: เปลี่ยน font ใน `_setup_styles()` เช่น `("Sarabun", 11)` |

---

## 🔒 Security Notes

- ไม่ commit `config.ini` ที่มี path ของข้อมูลภายในเข้า git
- ไม่ commit ไฟล์ Excel ที่มีข้อมูลจริงเข้า public repo
- หากใช้ใน private repo (PTT internal) ตามนโยบายปกติ
- ไฟล์ผลลัพธ์ใน `output/` ถูก gitignore ไว้แล้ว

---

## 📄 License

MIT License — ดูรายละเอียดใน [LICENSE](LICENSE)

---

## 👥 Contributing

ภายในทีม: pull request หรือแจ้งปัญหาผ่าน GitHub Issues

---

## 📧 Contact

[ใส่ชื่อ / email ของผู้ดูแล]
