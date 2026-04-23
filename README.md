# ☕ Internet Cafe Manager

A lightweight, single-file desktop application for managing an internet café — built with **Python + Tkinter**.  
Track PC sessions, log expenses, handle bookings, and export daily reports to Excel, all from a dark-themed GUI that runs instantly on any Windows PC.

---

## ✨ Features

### 🖥️ Real-Time PC Status Grid
- Visual monitor-style grid for **15 PCs**
- Live countdown timers for each active session
- Color-coded status with animated glow effects:
  - 🟢 **Available** — PC is free
  - 🔴 **Occupied** — session is running
  - 🟡 **Warning** — less than 5 minutes remaining
  - 🔴 **Expired** — time is up, flashes "TIME'S UP!"
  - 🟠 **Open Session** — pay-later / open-ended session
  - ⚫ **Offline** — PC marked as shut down
- Right-click any PC to **Shutdown / Turn On**
- Click an occupied PC to quickly jump to its edit form

### ⏱️ Session Management
- **Fixed-duration sessions** with preset buttons: 30 min, 1 hr, 1:30, 2 hr, 2:30, 3 hr, 4 hr, 5 hr
- **Custom duration** entry (e.g. `1:15` for 1 hour 15 minutes)
- **Open sessions** (∞) for pay-later setups — close with a Time Out picker
- Automatic **amount calculation** based on configurable rates
- Manual amount override supported
- Optional **comment** field per session
- Edit or delete any record from the session table
- Double-click a row or use the **✏ Edit Selected** button

### 🔖 Bookings
- Reserve a PC for a future customer
- Pending bookings shown with a **BOOKED** badge on the PC grid
- Arrival **notification alerts** pop up when booking time is reached

### 💰 Expenses Tracker
- Log daily expenses (name + amount)
- Live summary card: **Earnings → Expenses → Remaining**
- Expenses are included in the daily Excel export

### 📊 Excel Export (Auto-Save)
- Data is automatically saved to a **monthly Excel workbook** (`MMM-YYYY.xlsx`) inside the `data/` folder
- Each day gets its own **sheet** (`DD-Mon`)
- Styled reports with dark-red session rows and dark-blue expense rows
- Totals row shows **Total Earnings**, **Total Expenses**, and **Net Remaining**
- Opens with one click via the **📂 Excel Folder** button

### 🤖 AI Assistant
- Built-in AI chat panel powered by **OpenRouter API**
- Toggle with the **🤖 AI Assistant** button in the header
- Helps with session queries, customer suggestions, and general café management tasks
- Configure your free OpenRouter API key from Settings

### 🔔 Smart Notifications
- Rotating notification bar for time-sensitive alerts:
  - Session expiry warnings (5 minutes before)
  - Session expired alerts
  - Booking arrival alerts

### ⚙️ Settings
- Configurable price per duration preset (₱ amounts)
- OpenRouter API key management
- Settings persist in `config.json`

---

## 📁 Project Structure

```
Internet-Cafe-Manager/
├── cafe_manager.py     # Main application (single file)
├── config.json         # Saved rate settings & API key
├── requirements.txt    # Python dependencies
├── run.bat             # One-click launcher (Windows)
└── data/
    ├── cafe_data.json  # Today's in-memory session state
    └── MMM-YYYY.xlsx   # Monthly Excel report (auto-generated)
```

---

## 🚀 Getting Started

### Prerequisites
- Python 3.9 or later
- pip

### Installation

```bash
# 1. Clone the repository
git clone https://github.com/your-username/Internet-Cafe-Manager.git
cd Internet-Cafe-Manager

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the application
python cafe_manager.py
```

### Windows Quick Launch
Double-click **`run.bat`** — it automatically starts the app and shows an error message if Python or openpyxl is missing.

---

## ⚙️ Configuration

Open **Settings** (⚙ button, top-right) to set your pricing:

| Duration | Default Rate |
|----------|-------------|
| 30 min   | ₱ 50        |
| 1 hour   | ₱ 90        |
| 1:30     | ₱ 140       |
| 2 hours  | ₱ 180       |
| 2:30     | ₱ 230       |
| 3 hours  | ₱ 270       |
| 4 hours  | ₱ 360       |
| 5 hours  | ₱ 450       |

Rates are saved to `config.json` and used for automatic amount calculation.

To use the **AI Assistant**, enter your free [OpenRouter](https://openrouter.ai) API key in Settings.

---

## 📦 Dependencies

| Package     | Version     | Purpose                     |
|-------------|-------------|-----------------------------|
| `openpyxl`  | ≥ 3.1.0     | Excel report generation     |
| `tkinter`   | built-in    | Desktop GUI framework       |

Install all dependencies:
```bash
pip install -r requirements.txt
```

---

## 🗂️ Data & Persistence

- **Session state** is saved in `data/cafe_data.json` and automatically resets at the start of each new day.
- **Excel reports** accumulate across days — each day is a new sheet in the current month's workbook.
- Data is stored **locally only** — no internet connection required (except for the optional AI feature).

---

## 📋 Tab Overview

| Tab | Description |
|-----|-------------|
| ▶ Current Users | Active sessions only (not yet closed) |
| ☰ All Records | Full log of all today's sessions |
| 💰 Expenses | Daily expense log and earnings summary |
| 🔖 Bookings | Upcoming/pending PC reservations |

---

## 🛡️ Notes

- The app is designed for **single-operator use** on a Windows desktop.
- All data is stored locally; no database or server is needed.
- The `config.json` file contains your OpenRouter API key — keep it private and do **not** commit it to public repositories.

---

## 📄 License

This project is open source. Feel free to fork and adapt it for your own café or business needs.
