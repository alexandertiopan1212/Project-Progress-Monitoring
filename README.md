# 🏗️ Project Progress Monitoring (Google Sheets Automation)

An all-in-one project monitoring system using Google Sheets and Apps Script — designed for construction projects, field teams, and operations managers. This solution helps track critical path milestones, generate Gantt charts, document daily reports, and upload issue photos, all inside your spreadsheet workspace.

---

## 📋 Ready-to-Use Templates

- 🔗 **Main Project Monitoring Template**  
  [👉 Click here to open and copy](https://docs.google.com/spreadsheets/d/1jqRDPFWOEaCMW6N2xM1DZeTFyM0DwVQ33iACwOj6CFo/edit?usp=drive_link)

- 🔗 **Daily Report Template (required for daily cloning)**  
  [👉 Click here to open and copy](https://docs.google.com/spreadsheets/d/1cr5zfBkn9bysM8iCcGIan-sMYrT6BTNk32_u-GaVd-A/edit?usp=sharing)

> Make a copy of both templates into your own Google Drive before using.

---

## ⚙️ Script Button Assignments

To ensure full functionality, make sure the following buttons in the **Main Template** are assigned to the correct script functions:

| Sheet | Button Label             | Assigned Function            |
|-------|--------------------------|------------------------------|
| `Master` | *(Auto runs on open/edit)* | `generateGantt()` & `markKeyMilestones()` |
| `Master` | *(Optional Time Trigger)* | `createTimeTrigger()`       |
| `Master` | *(Delete Triggers)*   | `deleteAllTriggers()`       |
| `Master` | *(Daily Report Generator)* | `createDailyReport()`    |
| `Issues` | `Upload Photo`        | `showIssuePhotoUploader()`  |

Inside the **Daily Report Sheet (auto-generated)**, script-based modals will handle:

- ✔️ Select Today's Activities → `showActivitySelector()`
- ➡️ Select Next Day Activities → `insertNextDayActivities()`

---

## 🚀 Features Overview

### 📅 Smart Daily Report System
- Pick a report date using calendar popup.
- Automatically clones a blank daily report sheet from your linked **template**.
- Activities can be selected from the `Master` sheet via dynamic modals.

### 🖼️ Issue Photo Upload to Google Drive
- Upload image proof directly to a specified Google Drive folder.
- Generates a clickable hyperlink (`View`) in the `Photos` column of the `Issues` sheet.

### 📊 Weekly Gantt Chart with Auto Detection
- Visualize planning vs actual progress in color-coded Gantt bars.
- Auto-calculate parent dates from children.
- Highlight "CRITICAL" paths automatically.

---

## 📂 File Structure

| File Name               | Description |
|------------------------|-------------|
| `Code.gs`              | Main backend logic and triggers |
| `DatePickerDialog.html` | Popup to select date for daily report |
| `ActivitySelector.html` | Modal to select today's activities |
| `NextDaySelector.html`  | Modal to select next day activities |
| `IssuePhotoUploader.html` | Sidebar to upload issue photos |
| `README.md`             | This documentation file |

---

## 📸 Interface Preview

![image](https://github.com/user-attachments/assets/b6e29d90-2b6f-4cbb-83ae-1b1a866f47a0)
![image](https://github.com/user-attachments/assets/ad778065-3a32-4308-b140-cddb64702945)


---

## 🧠 Advanced Automation

- **`onOpen()`**: Auto-runs Gantt generation and milestone detection.
- **`onEdit()`**: Real-time refresh on any planning/actual date edit.
- **`createTimeTrigger()`**: Optional 5-min interval trigger to refresh Gantt chart.
- **`createSheetAndPickActivities(date)`**: Clones new daily report sheet based on the date input.

---

## 🔐 Configuration Notes

- Make sure to update the following **script constants**:
  - `templateFileId` → Your **Daily Report Template Spreadsheet ID**
  - `folderId` → Google Drive **Folder ID** for storing issue photos

---

## 🪪 License

This project is open-source and free for both personal and professional use.  
Feel free to fork, contribute, or suggest improvements.

---

Built with ❤️ using Google Apps Script and Google Sheets for smarter, field-friendly project tracking.
