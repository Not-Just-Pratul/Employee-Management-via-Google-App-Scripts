# Employee Management — Google Apps Script HRMS

A lightweight, spreadsheet-backed HRMS (Human Resource Management System) built with Google Apps Script and a responsive single-file frontend. It provides employee profiles, attendance with geofenced check-in/out, leave and gate-pass workflows, document & policy management, payroll lookup, and manager/admin dashboards — all driven by a Google Sheet.

**Demo / Deployment**

- **Web App URL:** https://script.google.com/macros/s/AKfycbw1f7Dz2Kfxqhan9XgWmILC1MtwbKH4yinejR2Ss63UW-TblMbmEVdLcMemCnNN2UjZ/exec

**Features**

- **Attendance**: Geofence-based check-in/check-out with late/half-day detection and working hours calculation.
- **Authentication**: Employee login (supports migrating plaintext passwords to SHA-256 hashed values).
- **Leave Management**: Request, approve/reject, and automatic leave-balance update on approval.
- **Gate Pass**: Create and manage gate-pass requests with approval workflow.
- **Employee Directory & Profiles**: View and edit official details; manager views for teams.
- **Documents & Policies**: Upload/download employee documents and company policies via Google Drive (proxying for mobile access).
- **Payroll Lookup**: Read-only salary structure lookup per employee.
- **Admin / Manager Dashboards**: Team summaries, pending approvals, and basic analytics.

**Tech Stack**

- Google Apps Script (V8 runtime)
- Google Sheets (data backend)
- Google Drive (file storage)
- HTML/CSS frontend (single-file UI: [Index.html](Index.html))

**Quickstart**

1. Make a copy of the spreadsheet you want to use as the data store, or create a new one and add the required sheets (see "Sheet requirements" below).
2. Open Extensions → Apps Script in the spreadsheet (or create a standalone Apps Script project and set `SPREADSHEET_ID`).
3. Add the project files from this repo into the Apps Script project (see: [Code.gs](Code.gs) and [Index.html](Index.html)).
4. In the Apps Script editor run the `setup()` function once (select `setup` and click ▶ Run). This saves the active spreadsheet ID into the script properties so the Web App uses the correct sheet.
5. Deploy as a Web App (Deploy → New deployment → select "Web app"). Recommended deployment settings:
	- **Execute as:** User deploying
	- **Who has access:** Anyone (even anonymous) — matches `appsscript.json` (`ANYONE_ANONYMOUS`).
6. Share the resulting deployment URL with users (see the Demo / Deployment link above).

**Sheet requirements**

The script expects (recommended) sheets with the following names. The code uses flexible header matching but these sheet names are referenced throughout:

- `Employee Master`
- `Login Details`
- `Role`
- `Attendance Sheet`
- `Leave Request`
- `Leave Balance`
- `Salary Structure`
- `Department Master`
- `Designation Master`
- `Manager Master`
- `Holiday List`
- `Announcements`
- `Policies`
- `Gate Pass`
- `Employee Documents`
- `Task Sheet`

For a smooth setup, seed headers and at least one row of sample data for `Login Details`, `Employee Master`, and `Leave Balance`.

**Configuration**

- `SPREADSHEET_ID`: If using a standalone Apps Script project, set this as a Script Property (Project Settings → Script Properties → Add `SPREADSHEET_ID`). Running `setup()` inside the desired spreadsheet writes this automatically.
- Default password for new employees: `Welcome@123` (the script hashes this value before storing in `Login Details`).
- Geofence & other constants live in `Code.gs` near the top (e.g., `OFFICE_LAT`, `OFFICE_LNG`, `OFFICE_RADIUS`, `DEFAULT_PASSWORD`). Update these values to match your office location and policy.

**Permissions / Scopes**

The project requires these scopes (declared in [appsscript.json](appsscript.json)):

- `https://www.googleapis.com/auth/spreadsheets`
- `https://www.googleapis.com/auth/script.container.ui`
- `https://www.googleapis.com/auth/userinfo.email`
- `https://www.googleapis.com/auth/drive`

When deploying for the first time, grant the requested permissions.

**Important notes & security**

- The web app in this repo is configured for `ANYONE_ANONYMOUS` access to allow mobile users to access public resources. If you need stricter access control, change the Web App access and adjust sharing settings accordingly.
- Passwords are migrated to SHA-256 hashed values on first successful login; existing plaintext passwords will be upgraded automatically.
- Files uploaded to Google Drive are set to "Anyone with the link can view" to simplify mobile access. Review your organization’s data policies before enabling this in production.

**Files of interest**

- [Code.gs](Code.gs) — Server-side Apps Script logic (authentication, attendance, leaves, documents, policies, admin functions).
- [Index.html](Index.html) — Frontend UI and client-side logic.
- [appsscript.json](appsscript.json) — Project configuration and OAuth scopes.

**Troubleshooting**

- If the script cannot find the spreadsheet, ensure `SPREADSHEET_ID` is set or run `setup()` from the desired sheet.
- Use `checkConfig()` (Apps Script editor → select `checkConfig` → Run) to verify `SPREADSHEET_ID` and sheet accessibility.

**Contributing**

Contributions, bug reports, and feature requests are welcome. Please open an issue or submit a PR.

**License**

This repository includes a `LICENSE` file — review it for license terms.

---
Generated from the project sources: [Code.gs](Code.gs) and [Index.html](Index.html).