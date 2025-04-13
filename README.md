# Python integrated Automatic WhatsApp Message Sender 📤

This project automates the process of sending customized messages via WhatsApp Desktop using Python. It reads data from an Excel sheet and sends messages to specified phone numbers using the WhatsApp desktop app.

---

## 📌 Features

- Reads recipient data from an Excel file
- Sends personalized messages using WhatsApp Desktop
- Formats salary/summary messages cleanly
- Includes current month and year dynamically
- Easy to customize message content
- Deployable as an executable (`.exe`) using PyInstaller

---

## 🛠️ Tech Stack

- **Language:** Python 3.x
 
- **Libraries:** 
  - `pandas` – for reading Excel files
  - `pyautogui` – for GUI automation
  - `pyperclip` – to copy/paste message text
  - `datetime` – for date formatting
  - `time` – for delays and synchronization
  
- **Deployment Tool:** : PyInstaller

---

## 📋 Excel File Format (`data.xlsx`)

The Excel sheet should contain the following columns:

| Name      | Phone Number | Days | Gross | Net | EPF | ESI | Advance | Mess | Store | Others |
|-----------|--------------|------|-------|-----|-----|-----|---------|------|-------|--------|
| John Doe  | 9123456789   | 28   | ...   | ... | ... | ... |   ...   | ...  |  ...  |  ...   |



