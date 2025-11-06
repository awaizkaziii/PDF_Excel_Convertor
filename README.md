
# 📄 PDF to Excel Converter & Excel Image OCR (Streamlit App)

A **comprehensive Streamlit application** that provides powerful PDF and Excel processing capabilities, including **PDF → Excel conversion** and **OCR text extraction** from images embedded in Excel files.

This app offers **two main tools**:

1. 🧾 **PDF to Excel Converter**
2. 🧠 **Text Extractor for Excel Images**

---

## 🚀 Features

### 🔹 1. PDF to Excel Converter

* Convert PDF files to Excel format using **Spire.PDF**
* Preserve layout and formatting (no structural changes)
* Supports multi-page PDFs

### 🔹 2. Text Extractor for Excel Images

* Extract text from **images embedded in Excel (.xlsx)** files using **Tesseract OCR**
* Specify which **columns** to process
* Works across **multiple sheets**
* Auto-detects Tesseract installation
* Flexible **text joining options** for multiple images in one cell
* Option to **remove images after OCR**, leaving clean text behind

---

## 🧩 Key Highlights

✅ Automatic Tesseract OCR detection
✅ Support for multi-sheet Excel files
✅ Specify columns for OCR processing
✅ Join extracted text flexibly (space, newline, or custom)
✅ Clean and ready-to-use Excel output

---

## ⚙️ Installation

### 1️⃣ Clone the Repository

```bash
git clone <repository-url>
cd <repository-name>
```

### 2️⃣ Install Required Dependencies

```bash
pip install -r requirements.txt
```

### 3️⃣ Install Tesseract OCR

#### 🪟 Windows

* Download from [Tesseract GitHub Releases](https://github.com/tesseract-ocr/tesseract/releases)
* Or install via Chocolatey:

  ```bash
  choco install tesseract
  ```

#### 🍏 macOS

```bash
brew install tesseract
```

#### 🐧 Linux (Ubuntu/Debian)

```bash
sudo apt-get update
sudo apt-get install tesseract-ocr
```

#### 🐧 Linux (CentOS/RHEL)

```bash
sudo yum install tesseract
```

---

## ▶️ Usage

Run the main application:

```bash
streamlit run app.py
```

Once launched, open your **web browser** — the app will start automatically.

### 🧾 PDF to Excel Converter

1. Upload a `.pdf` file
2. Click **Convert**
3. Download the resulting `.xlsx` file

### 🧠 Text Extractor for Excel Images

1. Upload an `.xlsx` file
2. Configure OCR settings in the sidebar
3. (Optional) Provide Tesseract path if not auto-detected
4. Specify target columns for OCR processing
5. Click **Run OCR**
6. Download the cleaned Excel file (images replaced with text)

---

## 📁 Project Structure

```
📦 Project
├── app.py                 # 🎯 Main Streamlit application
├── PDIG.py                # ⚙️ PDF → Excel converter (Spire.PDF)
├── Text_extractor.py      # 🧠 Excel Image OCR module
├── requirements.txt       # 📦 Dependencies
└── README.md              # 📝 This file
```

---

## 🧠 Dependencies

| Library       | Purpose                          |
| ------------- | -------------------------------- |
| `streamlit`   | Web app framework                |
| `pytesseract` | Python wrapper for Tesseract OCR |
| `openpyxl`    | Excel file manipulation          |
| `spire.pdf`   | PDF processing and conversion    |
| `Pillow`      | Image processing                 |
| `pymupdf`     | PDF → image conversion           |

> 🧩 **Tip:** Make sure all dependencies are installed and compatible with your Python version.

---

## 🧾 Excel Image OCR Workflow

1. Extract images from Excel cells
2. Process each image via **Tesseract OCR**
3. Replace image cells with extracted text
4. (Optional) Remove original images
5. Export a clean `.xlsx` file

---

## 🛠 Troubleshooting

### ❌ **Tesseract Not Found**

If not auto-detected:

* Verify Tesseract installation
* Set full path manually in the sidebar

**Common Paths**

* Windows → `C:\Program Files\Tesseract-OCR\tesseract.exe`
* macOS → `/usr/local/bin/tesseract`
* Linux → `/usr/bin/tesseract`

---

### ⚠️ **Conversion Issues**

* Ensure supported formats:

  * `.pdf` → for conversion
  * `.xlsx` → for OCR
* Check file permissions and size limits
* Verify all dependencies are installed correctly

---

## ❤️ Credits

Developed with **Python**, **Streamlit**, **Spire.PDF**, and **Tesseract OCR**
Empowering users to automate **PDF & Excel text extraction** with AI-powered precision.

---
