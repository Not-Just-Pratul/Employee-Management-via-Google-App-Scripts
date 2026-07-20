# Employee Management — Google Apps Script HRMS

A lightweight, spreadsheet-backed HRMS (Human Resource Management System) built with Google Apps Script and a responsive single-file frontend. It provides employee profiles, attendance with geofenced check-in/out, leave and gate-pass workflows, document & policy management, payroll lookup, and manager/admin dashboards — all driven by a Google Sheet.

---

## Demo / Deployment

* **Web App URL:**
  [https://script.google.com/macros/s/AKfycbw1f7Dz2Kfxqhan9XgWmILC1MtwbKH4yinejR2Ss63UW-TblMbmEVdLcMemCnNN2UjZ/exec](https://script.google.com/macros/s/AKfycbw1f7Dz2Kfxqhan9XgWmILC1MtwbKH4yinejR2Ss63UW-TblMbmEVdLcMemCnNN2UjZ/exec)

---

## Features

* **Attendance:** Geofence-based check-in/check-out with late/half-day detection and working hours calculation.
* **Authentication:** Employee login with automatic migration from plaintext passwords to SHA-256 hashed passwords.
* **Leave Management:** Submit, approve/reject leave requests with automatic leave balance updates.
* **Gate Pass:** Create and manage gate-pass requests with approval workflow.
* **Employee Directory & Profiles:** Employee profile management with manager/team views.
* **Documents & Policies:** Upload and download employee documents and company policies using Google Drive.
* **Payroll Lookup:** Read-only salary structure lookup.
* **Admin & Manager Dashboards:** Team summaries, pending approvals, and analytics.

---

## Tech Stack

* Google Apps Script (V8 Runtime)
* Google Sheets
* Google Drive
* HTML / CSS / JavaScript (Single Page UI)

---

# Quick Start

## Step 1 – Download the Repository

Clone or download this repository to your local machine.

---

## Step 2 – Create the Google Spreadsheet

Inside this repository, you will find an **Excel template**.

Create a new Google Spreadsheet and recreate the spreadsheet exactly according to the Excel template.

### **Important**

The following must remain exactly the same:

* Sheet names
* Column names
* Column order
* Overall sheet structure

Do **not** rename any sheet or modify headers unless you also update the Apps Script accordingly.

The HRMS uses these sheet names and headers internally, so changing them may cause features to stop working.

---

## Step 3 – Add Your Data

Once all sheets have been created successfully, you can replace the sample data with your own company information.

For example:

* Employee Details
* Login Details
* Leave Balance
* Departments
* Salary Structure
* Holidays
* Policies
* Documents

---

## Step 4 – Open Apps Script

From your Google Spreadsheet:

**Extensions → Apps Script**

---

## Step 5 – Create Project Files

Create the following files inside your Apps Script project.

```
Code.gs
Index.html
appsscript.json
```

Copy the contents of each file from this repository and paste them into the corresponding Apps Script files.

```
Repository
│
├── Code.gs
├── Index.html
└── appsscript.json
```

Save all files after pasting.

---

## Step 6 – Configure the Spreadsheet

Run the following function once:

```
setup()
```

This automatically saves the active Spreadsheet ID into the Script Properties.

If you are using a standalone Apps Script project instead of a container-bound project, manually add the following Script Property:

```
SPREADSHEET_ID
```

with the ID of your Google Spreadsheet.

---

## Step 7 – Deploy the Web App

Deploy the Apps Script project.

**Deploy → New Deployment → Web App**

Recommended settings:

**Execute as**

```
User deploying
```

**Who has access**

```
Anyone
```

or configure access according to your organization's security requirements.

---

## Step 8 – Done

After deployment, open the generated Web App URL.

Your HRMS application is now ready to use.

---

# Required Sheets

The spreadsheet should contain the following sheets.

```
Employee Master
Login Details
Role
Attendance Sheet
Leave Request
Leave Balance
Salary Structure
Department Master
Designation Master
Manager Master
Holiday List
Announcements
Policies
Gate Pass
Employee Documents
Task Sheet
```

---

# Configuration

The following values are configurable inside `Code.gs`.

* `SPREADSHEET_ID`
* `OFFICE_LAT`
* `OFFICE_LNG`
* `OFFICE_RADIUS`
* `DEFAULT_PASSWORD`

Default password for newly created employees:

```
Welcome@123
```

Passwords are automatically stored as SHA-256 hashes.

---

# Required OAuth Scopes

The project requires the following scopes:

```
https://www.googleapis.com/auth/spreadsheets
https://www.googleapis.com/auth/script.container.ui
https://www.googleapis.com/auth/userinfo.email
https://www.googleapis.com/auth/drive
```

Grant the requested permissions during the first deployment.

---

# Security Notes

* Passwords are automatically migrated to SHA-256 hashes after the first successful login.
* Uploaded files are stored in Google Drive.
* By default, uploaded files are shared using **Anyone with the link can view** for easier mobile access.
* Review your organization's security policy before using this configuration in production.
* If stronger authentication is required, modify the Web App deployment permissions.

---

# Project Structure

```
Repository
│
├── Code.gs                # Server-side logic
├── Index.html             # Frontend UI
├── appsscript.json        # Apps Script configuration
├── Employee Management.xlsx
├── LICENSE
└── README.md
```

---

# Troubleshooting

### Spreadsheet not found

Run:

```
setup()
```

or verify that the `SPREADSHEET_ID` Script Property is correctly configured.

---

### Configuration Check

Run:

```
checkConfig()
```

This verifies:

* Spreadsheet ID
* Spreadsheet accessibility
* Required sheets

---

# Contributing

Contributions, bug reports, and feature requests are welcome.

Feel free to open an Issue or submit a Pull Request.

---

# License

This repository includes a `LICENSE` file.

Please review it before using this project.

---

## Important

> **The Excel template included in this repository is the source of truth for the application.**
>
> Before running the project:
>
> * Create a Google Spreadsheet using the Excel template.
> * Keep every sheet name exactly the same.
> * Keep every column header exactly the same.
> * Keep the sheet order and structure unchanged.
> * Create `Code.gs`, `Index.html`, and `appsscript.json` in Apps Script and copy the code from this repository.
>
> Once everything has been copied and verified, deploy the project as a Web App. The application should then work without any additional configuration (except updating office location or other optional constants).

---

Generated from the project source files (`Code.gs`, `Index.html`, and `appsscript.json`).
