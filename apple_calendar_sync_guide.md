# 📱 Apple Calendar Sync & Automation Guide

This guide details three seamless methods to sync your daily work shifts from **Apple Calendar (iOS / macOS / iCloud)** directly into your **Work Hours & Wage Tracker Spreadsheet**.

---

## 🚀 Method 1: Automated Python iCal Parser (Recommended for PC & Mac)

Included in your project folder is `sync_apple_calendar.py`, a script that imports exported `.ics` calendar files directly into your Excel workbook.

### How it Works:
1. **Title Matching**:
   - Events titled with `Agra`, `Tandoori` → Automatically logged under **Agra Tandoori** (95.00 kr/hr).
   - Events titled with `Grand`, `Hotel` → Automatically logged under **Grand Hotel** (153.64 kr/hr).
   - Events titled with `Backup` → Automatically logged under **Backup Job** (120.00 kr/hr).
   - Any other shift → Logged under **Other / Special**.
2. **Unpaid Break Detection**:
   - Automatically detects notes in event descriptions like `break: 30`, `45 min break`, `rest: 15`. Defaults to **30 minutes** if unstated.
3. **Duplicate Prevention**:
   - Safely checks existing entries so re-running the sync will never duplicate a shift.

### How to Run:
1. **Export Calendar**:
   - **Mac**: Open Calendar → Select your Work calendar → `File` > `Export` > `Export...` saving as `Work.ics`.
   - **iCloud Web**: Log in to [iCloud.com Calendar](https://www.icloud.com/calendar), enable **Public Calendar**, copy the link, and save as `.ics`.
2. **Execute Sync**:
   Run the command in terminal:
   ```bash
   python C:\Users\gbrot\.gemini\antigravity\scratch\work_time_wage_tracker\sync_apple_calendar.py "C:\path\to\your\calendar.ics"
   ```

---

## 📲 Method 2: Apple Shortcuts App (1-Tap Sync on iPhone & Mac)

You can create an iOS / macOS **Shortcut** that extracts calendar events and appends them to your tracker automatically!

### Step-by-Step Setup:
1. Open the **Shortcuts** app on iPhone, iPad, or Mac.
2. Tap **+** to create a new Shortcut. Name it **"Sync Work Hours to Tracker"**.
3. Add the following actions:
   - **Find Calendar Events**:
     - *Calendar* is `Work`
     - *Start Date* is `in the last 30 days`
   - **Repeat with Each** item in *Calendar Events*:
     - **Get Details of Calendar Event**: Title, Start Date, End Date, Notes.
     - **Calculate Duration**: `(End Date - Start Date)` in Hours.
     - **Text Box**: Format row as `[Start Date],[Title],[Start Time],[End Time],[Break],[Hourly Rate]`
     - **Append to File**: Append the formatted text line to `Work_Shifts.csv` in iCloud Drive.
4. Add a button to your iPhone Home Screen or Widget! 1-tap syncs all new shifts.

---

## 🌐 Method 3: Live iCloud Webcal URL Subscription

If you want live real-time sync without manual exports:
1. Open **Apple Calendar** on Mac / iPhone.
2. Right-click your Work calendar and select **Share Calendar** -> Check **Public Calendar**.
3. Copy the URL (starts with `webcal://...`).
4. You can paste this Webcal URL into Google Sheets using `=IMPORTFEED()` / Google Apps Script or pass it directly to `sync_apple_calendar.py`:
   ```bash
   python sync_apple_calendar.py "webcal://p45-caldav.icloud.com/published/2/..."
   ```

---

## 🛠️ Summary of Files Created

| File | Description |
| :--- | :--- |
| **`Work_Hours_and_Wage_Tracker.xlsx`** | Complete multi-job tracker Excel workbook with Dashboard, Daily Log, Job Settings, Weekly/Monthly/Yearly Summaries. |
| **`generate_excel.py`** | Python script to recreate/rebuild the Excel file anytime. |
| **`sync_apple_calendar.py`** | Python sync utility for importing `.ics` Apple Calendar exports. |
| **`sample_apple_calendar.ics`** | Test calendar file pre-configured with Agra Tandoori, Grand Hotel, and Backup shifts. |
