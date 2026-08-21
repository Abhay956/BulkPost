# BulkPost — Bulk Email Sender with Attachments

A simple Python desktop app to send single or bulk emails via Gmail, with document attachment support (PDF, Word, Excel, images, etc.). Bulk mode reads recipient addresses from an Excel file.

## Requirements

- Python 3.8+
- Gmail account with an [App Password](https://myaccount.google.com/apppasswords)

## Setup

**Windows**
```powershell
git clone https://github.com/Abhay956/BulkPost.git
cd BulkPost
py -m pip install -r requirements.txt
python app.py
```

**Linux**
```bash
git clone https://github.com/Abhay956/BulkPost.git
cd BulkPost
sudo apt install python3-tk
pip install -r requirements.txt
python3 app.py
```

## Usage

1. Open Settings (gear icon) → enter your Gmail address and App Password → Save.
2. Choose **Single** (type one email) or **Multiple** (browse and select an `.xlsx` file with an `Email` column).
3. Enter Subject and Message.
4. (Optional) Click **Attach File** to add a document.
5. Click **Send**.
<img width="1002" height="650" alt="image" src="https://github.com/user-attachments/assets/970347da-0a34-4572-96eb-bf92546ebb8a" />


## Developer

Built by **Abhay Pande**
