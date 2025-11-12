# 💻 Smart Terminal Assistant

**Voice-controlled automation assistant for PDFs, Word, Emails, and File Management — directly from your terminal.**  
Built completely in Python 🐍, this tool lets you perform real-world productivity tasks faster — offline.

---

## 🚀 Features

✅ **PDF Tools**
- Merge multiple PDFs  
- Split PDFs by page range  
- Protect & unlock PDFs with passwords  
- Convert Images → PDF  

✅ **Word Tools**
- Convert Word → PDF  
- Merge or Split Word files automatically  

✅ **File Tools**
- Auto-organize files by type  
- Move or clean large folders  

✅ **Email Automation**
- Send multiple files directly via Gmail  
- Built-in SMTP support  

✅ **Voice Assistant**
- Hands-free commands like:  
  - “Merge PDF”  
  - “Convert Word to PDF”  
  - “Split Word File”  
  - “Exit”  

---

## 🧠 Built With

- `Python 3.10+`
- `SpeechRecognition` – Voice input  
- `gTTS` + `VLC` – Natural speech output  
- `PyPDF2`, `python-docx`, `Pillow` – File processing  
- `Rich` – Beautiful terminal interface  
- `smtplib` – Email handling  

---

## ⚙️ Installation (Ubuntu/Linux)

### 1️⃣ Clone the repository
```bash
git clone https://github.com/ASWINa1636/Smart-Terminal.git
cd Smart-Terminal

### 2️⃣ Create a virtual environment (recommended)
python3 -m venv venv
source venv/bin/activate

3️⃣ Install dependencies
pip install -r requirements.txt

4️⃣ Run the assistant
python3 main.py

🎙️ Voice Assistant Mode (Ubuntu)

Then simply say:

“Merge PDF”
“Convert Word to PDF”
“Exit”

🧩 Package Structure
smart_terminal/
│
├── main.py                     # CLI entry point
├── modules/
│   ├── pdf_tools.py
│   ├── file_organizer.py
│   ├── email_report.py
│   ├── voice_assistant.py
│   └── __init__.py
│
├── requirements.txt
└── README.md


🤝 Contributing

Pull requests are welcome!
If you’d like to contribute new features (like OCR, file compression, or email templates), fork the repo and submit a PR.

📜 License

This project is licensed under the MIT License

⭐ Support

If you like this project, give it a ⭐ on GitHub!
Your star helps motivate development of more open-source automation tools ❤️
