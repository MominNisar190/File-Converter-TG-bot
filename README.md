# 👑 GGOD FATHERR VCF Maker Bot

A premium Telegram bot for converting, merging, splitting, and renaming contact files.  
Supports **.txt · .csv · .vcf · .xlsx** as input for all features.

---

## ⚡ Features

| Feature | Description |
|---|---|
| 📁 Any → VCF | Convert any contact file to VCF |
| 📊 Any → CSV/XLSX | Export contacts to CSV or Excel |
| 📇 Any → VCF | Import from any format to VCF |
| 📄 Any → TXT | Extract phone numbers from any file |
| 💬 MSG → TXT | Save any message as a TXT file |
| 📝 Rename File | Rename any uploaded file |
| ✏️ Rename Contacts | Bulk rename VCF contact names with a prefix |
| 🧩 Merge Files | Merge multiple files into one |
| ✂️ Split File | Split a large file into smaller chunks |
| 🛡️ Navy Format | Apply admin navy numbering to contacts |

---

## 🚀 Deploy on Render

### 1. Push to GitHub
Make sure your code is pushed to a GitHub repository.

### 2. Create a new Render service
1. Go to [render.com](https://render.com) → **New** → **Background Worker**
2. Connect your GitHub repo
3. Render will auto-detect `render.yaml`

### 3. Set Environment Variables
In Render dashboard → **Environment** tab, add:

| Key | Value |
|---|---|
| `BOT_TOKEN` | Your Telegram bot token |
| `PREMIUM_CODE` | Your premium unlock code (e.g. `#000#`) |

### 4. Deploy
Click **Deploy** — the bot will start automatically.

---

## 💻 Run Locally

```bash
pip install -r requirements.txt
BOT_TOKEN=your_token PREMIUM_CODE=#000# python bot.py
```

Or on Windows:
```bash
set BOT_TOKEN=your_token
set PREMIUM_CODE=#000#
python bot.py
```

---

## 🔐 Premium System

Users must enter a premium code to unlock all features.  
Contact admin to get access: [@ggod_fatherr](https://t.me/ggod_fatherr)

---

## 👤 Admin

Telegram: [@ggod_fatherr](https://t.me/ggod_fatherr)
