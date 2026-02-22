# Setup Guide

## 1. Google Apps Script Setup

Open Measure Template → Extensions → Apps Script.

Create files:

- Code.gs
- SidebarUI.gs
- sidebar.html

Paste repository code into matching files.

Save project.

---

## 2. Script Properties Required

Add via:

Apps Script → Project Settings → Script Properties

Required keys:

COMPANYCAM_TOKEN=
LP_API_BASE=
LP_API_USERNAME=
LP_API_PASSWORD=
LP_APPKEY=

---

## 3. Install Menu

Reload spreadsheet.

Use:

Extensions → Apps Script → Run:

setupMeasureTemplateConfig_

Authorize when prompted.

---

## 4. Sidebar

Open any measure packet.

Click custom menu or sidebar launcher.

Confirm sidebar loads.

---

## 5. Calendar Automation

Replace calendar automation script ONLY after:

- LP API credentials added.
- CompanyCam token verified.

---

## 6. Testing Checklist

- Create test measure.
- Confirm packet auto-creates.
- Confirm JobID written to C1.
- Confirm GrossAmount → D5.
- Confirm BalanceDue → G5.
- Confirm CompanyCam link inserted.
- Confirm PDF upload buttons appear.
