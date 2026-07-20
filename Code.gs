// ─── CONSTANTS ──────────────────────────────────────────────────────────────
const DEFAULT_PASSWORD = 'Welcome@123';
const CACHE_TTL = 30; // seconds

// ─── SPREADSHEET ACCESSOR (lazy, portable) ───────────────────────────────────
/**
 * Returns the bound spreadsheet.
 * Resolution order:
 *   1. Script Property "SPREADSHEET_ID"  — use this for standalone deployments
 *   2. SpreadsheetApp.getActiveSpreadsheet() — works for container-bound scripts
 * Throws a clear, actionable error if neither resolves.
 */
function _getSpreadsheet() {
  try {
    const props = PropertiesService.getScriptProperties();
    const ssId  = props.getProperty('SPREADSHEET_ID');
    if (ssId && ssId.trim()) {
      return SpreadsheetApp.openById(ssId.trim());
    }
  } catch(e) {
    // PropertiesService unavailable or property missing — fall through
  }
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch(e) {
    // Not container-bound — fall through
  }
  throw new Error(
    'Spreadsheet not found. Please set the SPREADSHEET_ID script property ' +
    '(Project Settings → Script Properties → Add property: SPREADSHEET_ID = <your sheet ID>).'
  );
}

// ─── HELPERS ────────────────────────────────────────────────────────────────

/** Get sheet with error handling */
function _getSheet(name) {
  const sheet = _getSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" not found.');
  return sheet;
}

// ─── ONE-TIME SETUP HELPER ───────────────────────────────────────────────────
/**
 * Run this function ONCE from the Apps Script editor after copying the script
 * to a new account.  It saves the active spreadsheet's ID as a script property
 * so the Web App always opens the correct sheet regardless of which Google
 * account is running it.
 *
 * Steps:
 *   1. Open the spreadsheet you want to use.
 *   2. Open Extensions → Apps Script.
 *   3. Select the function "setup" in the toolbar and click ▶ Run.
 *   4. Grant permissions when prompted.
 *   5. Deploy as Web App — done.
 */
function setup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log('ERROR: No active spreadsheet found. Open the spreadsheet first, then run setup().');
      return;
    }
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
    Logger.log('✅ Setup complete! SPREADSHEET_ID saved: ' + ss.getId());
    Logger.log('   Spreadsheet name: ' + ss.getName());
    Logger.log('   You can now deploy this script as a Web App.');
  } catch(e) {
    Logger.log('ERROR during setup: ' + e.message);
  }
}

/**
 * Utility: show the currently configured spreadsheet ID in the logs.
 * Run from the Apps Script editor to verify the configuration.
 */
function checkConfig() {
  try {
    const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (ssId) {
      Logger.log('SPREADSHEET_ID is set: ' + ssId);
      try {
        const ss = SpreadsheetApp.openById(ssId);
        Logger.log('✅ Spreadsheet found: ' + ss.getName());
      } catch(e) {
        Logger.log('❌ Cannot open spreadsheet: ' + e.message);
      }
    } else {
      Logger.log('SPREADSHEET_ID is NOT set. Will use getActiveSpreadsheet() as fallback.');
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        Logger.log('Active spreadsheet: ' + (ss ? ss.getName() : 'none'));
      } catch(e) {
        Logger.log('No active spreadsheet either. Run setup() first.');
      }
    }
  } catch(e) {
    Logger.log('ERROR: ' + e.message);
  }
}

/** Get sheet data with caching */
function _sheetData(name) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'sheet_' + name;
  const cached = cache.get(cacheKey);
  
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) {
      // Cache corrupted, fetch fresh
    }
  }
  
  const data = _getSheet(name).getDataRange().getValues();
  
  // Cache for read-heavy sheets only
  const cacheable = ['Employee Master', 'Holiday List', 'Department Master', 'Designation Master', 'Manager Master', 'Announcements'];
  if (cacheable.indexOf(name) !== -1) {
    try {
      cache.put(cacheKey, JSON.stringify(data), CACHE_TTL);
    } catch(e) {
      // Cache too large, skip caching
    }
  }
  
  return data;
}

/** Invalidate cache for a sheet after write operations */
function _invalidateCache(sheetName) {
  const cache = CacheService.getScriptCache();
  cache.remove('sheet_' + sheetName);
}

/** Simple password hashing using SHA-256 */
function _hashPassword(password) {
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return Utilities.base64Encode(hash);
}

/** Acquire lock for write operations */
function _acquireLock() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Wait up to 10 seconds
    return lock;
  } catch(e) {
    throw new Error('Could not acquire lock. Please try again.');
  }
}

/** Release lock */
function _releaseLock(lock) {
  if (lock) {
    try {
      lock.releaseLock();
    } catch(e) {}
  }
}

function _formatDate(d) {
  if (!d) return '';
  if (d instanceof Date) {
    // Guard against epoch-zero dates (time-only cells read as Date)
    if (d.getFullYear() < 1900) return '';
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  const s = String(d).trim();
  if (!s) return '';

  // ISO 8601 string: "2026-02-01T18:30:00.000Z" or "2026-02-01"
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const parsed = new Date(s);
    if (!isNaN(parsed.getTime()) && parsed.getFullYear() >= 1900) {
      return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
  }

  return s;
}

function _formatTime(t) {
  if (!t) return '--';
  if (t instanceof Date) {
    return Utilities.formatDate(t, Session.getScriptTimeZone(), 'HH:mm:ss');
  }
  // Already a string like "13:05:19"
  return String(t);
}

function _today() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function _now() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');
}



/**
 * Flexible column index finder — tries exact match first, then
 * falls back to case-insensitive + underscore/space-normalised match.
 * This prevents silent failures when sheet headers differ slightly.
 */
function _colIdx(headers, name) {
  // 1. Exact match
  const exact = headers.indexOf(name);
  if (exact !== -1) return exact;
  // 2. Normalised match: lowercase, replace underscores/hyphens/spaces with nothing
  const norm = s => String(s).toLowerCase().replace(/[\s_\-]/g, '');
  const target = norm(name);
  for (let i = 0; i < headers.length; i++) {
    if (norm(headers[i]) === target) return i;
  }
  return -1;
}

/**
 * Convert any Google Drive sharing URL to a direct thumbnail URL
 * so it can be used in <img src="..."> without CORS issues.
 * Non-Drive URLs are returned as-is.
 */
function _drivePhotoUrl(url) {
  if (!url) return '';
  const s = String(url).trim();
  // Pattern: https://drive.google.com/file/d/FILE_ID/view?...
  const m1 = s.match(/drive\.google\.com\/file\/d\/([^\/\?]+)/);
  if (m1) return 'https://drive.google.com/uc?export=view&id=' + m1[1];
  // Pattern: https://drive.google.com/open?id=FILE_ID
  const m2 = s.match(/drive\.google\.com\/open\?id=([^&]+)/);
  if (m2) return 'https://drive.google.com/uc?export=view&id=' + m2[1];
  // Pattern: https://drive.google.com/uc?id=FILE_ID or uc?export=view&id=FILE_ID or thumbnail?id=FILE_ID
  const m3 = s.match(/drive\.google\.com\/(?:uc|thumbnail)\?.*id=([^&]+)/);
  if (m3) return 'https://drive.google.com/uc?export=view&id=' + m3[1];
  return s;
}

// ─── ROUTING ────────────────────────────────────────────────────────────────

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Workforce Operations Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── AUTHENTICATION ──────────────────────────────────────────────────────────

function loginUser(empCode, password) {
  try {
    const loginData = _sheetData("Login Details");
    let authUser = null;
    const hashedInput = _hashPassword(password);

    // Match by Employee Code (col 0) and Password (col 3)
    // Supports both plain-text (legacy) and SHA-256 hashed passwords
    for (let i = 1; i < loginData.length; i++) {
      if (String(loginData[i][0]).trim().toUpperCase() !== String(empCode).trim().toUpperCase()) continue;
      const storedPwd = String(loginData[i][3]);
      const isMatch = storedPwd === password || storedPwd === hashedInput;
      if (isMatch) {
        authUser = {
          empCode: String(loginData[i][0]),
          name:    loginData[i][1],
          email:   loginData[i][2]
        };
        // Migrate plain-text password to hashed on first login
        if (storedPwd === password && storedPwd !== hashedInput) {
          try {
            const sheet = _getSheet("Login Details");
            sheet.getRange(i + 1, 4).setValue(hashedInput);
          } catch(e) {}
        }
        break;
      }
    }
    if (!authUser) return { success: false, message: "Invalid employee code or password." };

    // Auto-detect highest role from Role sheet
    let detectedRole = 'Employee';
    try {
      const roleData = _sheetData("Role");
      const rolePriority = { 'Super Admin': 4, 'Admin': 3, 'Manager': 2, 'Employee': 1 };
      for (let j = 1; j < roleData.length; j++) {
        const rowEmp  = String(roleData[j][2] || '');
        const rowRole = String(roleData[j][1] || '');
        if (rowEmp === authUser.empCode) {
          if ((rolePriority[rowRole] || 0) > (rolePriority[detectedRole] || 0)) {
            detectedRole = rowRole;
          }
        }
      }
    } catch(e) {}

    // Write audit log
    _writeAuditLog(authUser.empCode, 'LOGIN', 'User logged in', '');

    return { success: true, user: { ...authUser, role: detectedRole } };
  } catch (e) {
    return { success: false, message: "Login error: " + e.message };
  }
}

// ─── PROFILE ─────────────────────────────────────────────────────────────────

function getFullProfile(empCode) {
  const data = _sheetData("Employee Master");
  const h = data[0];

    const photoColRaw = _colIdx(h, "Emp Photo URL");
    const idx = {
      id:          _colIdx(h, "Emp ID"),
      name:        _colIdx(h, "Employee Name"),
      empRole:     _colIdx(h, "Employment Role"),
      dob:         _colIdx(h, "DOB"),
      aadhaar:     _colIdx(h, "Aadhar Number"),
      pan:         _colIdx(h, "PAN Number"),
      gender:      _colIdx(h, "Gender"),
      persMob:     _colIdx(h, "Personal Mobile Number"),
      offMob:      _colIdx(h, "Office Mobile Number"),
      persEmail:   _colIdx(h, "Personal Email ID"),
      offEmail:    _colIdx(h, "Official Email ID"),
      city:        _colIdx(h, "City"),
      district:    _colIdx(h, "District"),
      state:       _colIdx(h, "State"),
      resignMonth: _colIdx(h, "Resign Month"),
      level:       _colIdx(h, "Level"),
      process:     _colIdx(h, "Process"),
      account:     _colIdx(h, "Account Number"),
      ifsc:        _colIdx(h, "IFSC Code"),
      shift:       _colIdx(h, "Shift ID"),
      status:      _colIdx(h, "Status"),
      mgr:         _colIdx(h, "Manager ID"),
      address:     _colIdx(h, "Permanent Address"),
      photo:       photoColRaw >= 0 ? photoColRaw : 23,  // col X fallback
      doj:         _colIdx(h, "Date of Joining")
    };

  // Build a quick empCode→name lookup for manager name resolution
  const empNameMap = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][idx.id]) empNameMap[String(data[i][idx.id])] = data[i][idx.name];
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.id]) === String(empCode)) {
      const mgrId   = idx.mgr >= 0 ? String(data[i][idx.mgr] || '') : '';
      const mgrName = empNameMap[mgrId] || mgrId || '—';
      return {
        empCode:      String(data[i][idx.id]),
        name:         data[i][idx.name],
        // Personal fields
        personalMobile: idx.persMob   >= 0 ? String(data[i][idx.persMob]   || '') : '',
        personalEmail:  idx.persEmail >= 0 ? String(data[i][idx.persEmail] || '') : '',
        aadhaar:        idx.aadhaar   >= 0 ? String(data[i][idx.aadhaar]   || '') : '',
        pan:            idx.pan       >= 0 ? String(data[i][idx.pan]       || '') : '',
        gender:         idx.gender    >= 0 ? String(data[i][idx.gender]    || '') : '',
        dob:            _formatDate(idx.dob >= 0 ? data[i][idx.dob] : ''),
        address:        idx.address   >= 0 ? String(data[i][idx.address]   || '') : '',
        city:           idx.city      >= 0 ? String(data[i][idx.city]      || '') : '',
        district:       idx.district  >= 0 ? String(data[i][idx.district]  || '') : '',
        state:          idx.state     >= 0 ? String(data[i][idx.state]     || '') : '',
        // Official fields
        officialMobile: idx.offMob    >= 0 ? String(data[i][idx.offMob]    || '') : '',
        officialEmail:  idx.offEmail  >= 0 ? String(data[i][idx.offEmail]  || '') : '',
        employmentRole: idx.empRole   >= 0 ? String(data[i][idx.empRole]   || '') : '',
        level:          idx.level     >= 0 ? String(data[i][idx.level]     || '') : '',
        process:        idx.process   >= 0 ? String(data[i][idx.process]   || '') : '',
        dept:           idx.level     >= 0 ? String(data[i][idx.level]     || '') : '', // dept mapped to level
        designation:    idx.level     >= 0 ? String(data[i][idx.level]     || '') : '', // designation mapped to level
        accountNumber:  idx.account   >= 0 ? String(data[i][idx.account]   || '') : '',
        ifscCode:       idx.ifsc      >= 0 ? String(data[i][idx.ifsc]      || '') : '',
        resignMonth:    idx.resignMonth >= 0 ? String(data[i][idx.resignMonth] || '') : '',
        joiningDate:    _formatDate(idx.doj >= 0 ? data[i][idx.doj] : ''),
        shiftId:        idx.shift     >= 0 ? String(data[i][idx.shift]     || '') : '',
        status:         idx.status    >= 0 ? (String(data[i][idx.status] || '') || 'Active') : 'Active',
        managerId:      mgrId,
        managerName:    mgrName,
        role:           '', // Role now determined from Role sheet, not Employee Master
        photoUrl:       _drivePhotoUrl(String(data[i][idx.photo] || ''))
      };
    }
  }
  throw new Error("Profile not found for: " + empCode);
}

// ─── DASHBOARD ───────────────────────────────────────────────────────────────

function getDashboardData(empCode, role) {
  try {
    const today = _today();
    const attData = _sheetData("Attendance Sheet");
    const empData = _sheetData("Employee Master");

    // Today's attendance for this employee
    let todayIn = '', todayOut = '', todayStatus = '';
    for (let i = attData.length - 1; i >= 1; i--) {
      if (String(attData[i][1]).trim() !== String(empCode).trim()) continue;
      const rawDate = attData[i][2];
      const rowDateStr = rawDate instanceof Date
        ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'dd/MM/yyyy')
        : String(rawDate).trim();
      if (rowDateStr === today) {
        todayIn     = attData[i][3] ? _formatTime(attData[i][3]) : '';
        todayOut    = attData[i][4] ? _formatTime(attData[i][4]) : '';
        todayStatus = attData[i][7] || (attData[i][3] ? 'Present' : '');
        break;
      }
    }

    // Monthly stats for this employee
    const now = new Date();
    const monthStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MM/yyyy');
    let present = 0, absent = 0, late = 0, halfDay = 0;

    for (let i = 1; i < attData.length; i++) {
      if (String(attData[i][1]).trim() !== String(empCode).trim()) continue;

      let rowMonthYear = '';
      const rawDate = attData[i][2];
      if (rawDate instanceof Date) {
        if (rawDate.getFullYear() >= 1900) {
          rowMonthYear = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'MM/yyyy');
        }
      } else {
        const parts = String(rawDate).trim().split('/');
        if (parts.length === 3) {
          rowMonthYear = parts[1].padStart(2, '0') + '/' + parts[2];
        } else if (parts.length === 2) {
          rowMonthYear = String(rawDate).trim();
        }
      }

      if (rowMonthYear !== monthStr) continue;

      const st       = String(attData[i][7] || '').toLowerCase().replace(/\s+/g, ' ').trim();
      const markedAs = String(attData[i][8] || '').toLowerCase().replace(/\s+/g, ' ').trim();
      if (st === 'absent') {
        absent++;
      } else if (st === 'leave') {
        // leave — not counted as present or absent
      } else if (markedAs.indexOf('half') !== -1) {
        halfDay++;
        present++;
      } else if (markedAs.indexOf('late') !== -1) {
        present++; late++;
      } else if (attData[i][3]) {
        present++;
      }
    }

    // Leave balance
    let leaveBalance = { cl: 0, sl: 0, pl: 0 };
    try {
      const lbData = _sheetData("Leave Balance");
      for (let i = 1; i < lbData.length; i++) {
        if (String(lbData[i][0]).trim() === String(empCode).trim()) {
          leaveBalance = {
            cl: Number(lbData[i][1]) || 0,
            sl: Number(lbData[i][2]) || 0,
            pl: Number(lbData[i][3]) || 0
          };
          break;
        }
      }
    } catch(e) {}

    // Admin/Manager extras
    let totalEmp = 0, activeEmp = 0, pendingLeaves = 0;
    if (role === 'Admin' || role === 'Super Admin' || role === 'Manager') {
      const empH = empData[0];
      const statColIdx = _colIdx(empH, "Status");
      for (let i = 1; i < empData.length; i++) {
        if (empData[i][0]) {
          totalEmp++;
          const statusVal = statColIdx >= 0 ? String(empData[i][statColIdx] || '') : '';
          if (statusVal.toLowerCase() === 'active') activeEmp++;
        }
      }
      try {
        const lrData = _sheetData("Leave Request");
        for (let i = 1; i < lrData.length; i++) {
          if (String(lrData[i][8] || '').toLowerCase() === 'pending') {
            if (role === 'Manager') {
              if (String(lrData[i][10]) === String(empCode)) pendingLeaves++;
            } else {
              pendingLeaves++;
            }
          }
        }
      } catch(e) {}
    }

    return {
      today, todayIn, todayOut, todayStatus,
      monthly: { present, absent, late, halfDay },
      leaveBalance,
      admin: { totalEmp, activeEmp, pendingLeaves }
    };
  } catch(e) {
    return { error: e.message };
  }
}

// ─── ATTENDANCE ───────────────────────────────────────────────────────────────

// ─── GEOFENCE CONFIG ─────────────────────────────────────────────────────────
const OFFICE_LAT    = 28.2089;
const OFFICE_LNG    = 76.8601;
const OFFICE_RADIUS = 200; // metres

/**
 * Haversine distance between two lat/lng points (returns metres).
 */
function _haversineDistance(lat1, lng1, lat2, lng2) {
  const R   = 6371000; // Earth radius in metres
  const toR = Math.PI / 180;
  const dLat = (lat2 - lat1) * toR;
  const dLng = (lng2 - lng1) * toR;
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2)
          + Math.cos(lat1 * toR) * Math.cos(lat2 * toR)
          * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function markAttendance(empId, type, lat, lng) {
  try {
    // Location is mandatory — reject if not provided
    if (!lat || !lng) {
      return { status: 'error', msg: 'Location is required to mark attendance. Please enable location services and try again.' };
    }

    // Geofence check — must be within OFFICE_RADIUS metres of the office
    const distance = Math.round(_haversineDistance(Number(lat), Number(lng), OFFICE_LAT, OFFICE_LNG));
    if (distance > OFFICE_RADIUS) {
      return {
        status: 'error',
        msg: 'You are ' + distance + ' m away from Rishi Seals Pvt. Ltd. Check-in is only allowed within ' + OFFICE_RADIUS + ' m of the office premises.'
      };
    }
    const availability = 'Check-In at Rishi Seals Pvt. Ltd';

    const sheet = _getSheet("Attendance Sheet");
    const today = _today();
    const time  = _now();
    const loc   = lat + ',' + lng;

    if (type === 'IN') {
      // Prevent double check-in
      const data = sheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][1]) === String(empId) && _formatDate(data[i][2]) === today) {
          return { status: 'error', msg: 'Already checked in today at ' + data[i][3] };
        }
      }
      
      // Get employee's shift ID from Employee Master
      let shiftId = '';
      try {
        const empData = _sheetData("Employee Master");
        const h = empData[0];
        const idCol    = _colIdx(h, "Emp ID");
        const shiftCol = _colIdx(h, "Shift ID");   // flexible match — handles "Shift ID", "shift_id", etc.
        for (let i = 1; i < empData.length; i++) {
          if (String(empData[i][idCol]) === String(empId)) {
            shiftId = shiftCol >= 0 ? String(empData[i][shiftCol] || '').trim() : '';
            break;
          }
        }
      } catch(e) {}
      
      // Determine attendance status based on shift and check-in time
      let status = 'Present';
      const checkInTime = time; // Format: HH:mm:ss
      const timeParts = checkInTime.split(':');
      const checkInMinutes = parseInt(timeParts[0]) * 60 + parseInt(timeParts[1]);

      if (shiftId === 'SH_02') {
        // Night Shift: Start 20:00, Late Mark 20:30, Half Day 01:30 (next day)
        const lateMarkMinutes = 20 * 60 + 30;  // 20:30
        const halfDayMinutes  = 1 * 60 + 30;   // 01:30
        
        // Night shift spans midnight — check-in can be evening (20:00+) or early morning (00:00–04:30)
        if (checkInMinutes >= 0 && checkInMinutes <= 270) {
          // Early morning check-in (00:00 to 04:30)
          if (checkInMinutes > halfDayMinutes) status = 'Half Day';
        } else if (checkInMinutes >= 1200) {
          // Evening check-in (20:00 onwards)
          if (checkInMinutes > lateMarkMinutes) status = 'Late';
        }
      } else {
        // SH_01 (Day Shift), unassigned, or any other shift — default office hours
        // Start 09:00 | Late after 09:30 | Half Day after 13:30
        const lateMarkMinutes = 9 * 60 + 30;   // 09:30
        const halfDayMinutes  = 13 * 60 + 30;  // 13:30
        if (checkInMinutes > halfDayMinutes) {
          status = 'Half Day';
        } else if (checkInMinutes > lateMarkMinutes) {
          status = 'Late';
        }
      }
      
      // Determine Marked_As (punctuality) and Status (attendance)
      // Status   → always: Present / Absent / Leave / Gate Pass / Holiday / Week Off
      // Marked_As → On Time / Late Arrival / Half Day (punctuality detail)
      let markedAs = 'On Time'; // I: Marked_As — punctuality detail
      if (status === 'Half Day') {
        markedAs = 'Half Day';
      } else if (status === 'Late') {
        markedAs = 'Late Arrival';
      }

      const attId = 'ATT' + new Date().getTime();
      // Columns: A=AttID | B=EmpID | C=Date | D=CheckIn | E=CheckOut | F=Location | G=Availability | H=Status | I=Marked_As | J=Working_Hours
      sheet.appendRow([attId, empId, today, time, '', loc, availability, 'Present', markedAs, '']);
      
      let msg = 'Checked in at ' + time;
      if (status === 'Late')     msg += ' — You have arrived late.';
      if (status === 'Half Day') msg += ' — Marked as Half Day.';
      
      return { status: 'success', msg: msg, checkIn: time, markedAs: markedAs };

    } else {
      const data = sheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][1]) === String(empId) && _formatDate(data[i][2]) === today && !data[i][4]) {
          // Save checkout time
          sheet.getRange(i + 1, 5).setValue(time);

          // Calculate working hours from check-in
          let workingHours = '';
          try {
            const checkInRaw = data[i][3];
            const checkInStr = checkInRaw instanceof Date
              ? Utilities.formatDate(checkInRaw, Session.getScriptTimeZone(), 'HH:mm:ss')
              : String(checkInRaw);
            const [inH, inM, inS] = checkInStr.split(':').map(Number);
            const [outH, outM, outS] = time.split(':').map(Number);
            let totalSecs = (outH * 3600 + outM * 60 + outS) - (inH * 3600 + inM * 60 + inS);
            // Handle overnight shifts (checkout next day)
            if (totalSecs < 0) totalSecs += 24 * 3600;
            const hrs  = Math.floor(totalSecs / 3600);
            const mins = Math.floor((totalSecs % 3600) / 60);
            workingHours = hrs + 'h ' + String(mins).padStart(2, '0') + 'm';
          } catch(e) {}

          // Write Working_Hours to col J (1-based col 10)
          sheet.getRange(i + 1, 10).setValue(workingHours);

          return { status: 'success', msg: 'Checked out at ' + time + (workingHours ? ' · Working Hours: ' + workingHours : ''), checkOut: time, workingHours: workingHours };
        }
      }
      return { status: 'error', msg: 'No active check-in found for today.' };
    }
  } catch(e) {
    return { status: 'error', msg: e.message };
  }
}

function getAttendanceHistory(empCode, monthOffset) {
  try {
    const data = _sheetData("Attendance Sheet");
    const now  = new Date();
    now.setMonth(now.getMonth() - (monthOffset || 0));
    const targetMonth = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MM/yyyy');
    const records = [];

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() !== String(empCode).trim()) continue;
      
      let rowMonthYear = '';
      const rawDate = data[i][2];
      if (rawDate instanceof Date) {
        rowMonthYear = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'MM/yyyy');
      } else {
        // Handle string dates: dd/MM/yyyy or d/M/yyyy
        const parts = String(rawDate).trim().split('/');
        if (parts.length === 3) {
          // parts[1] = month, parts[2] = year
          const m = parts[1].padStart(2, '0');
          const y = parts[2];
          rowMonthYear = m + '/' + y;
        } else if (parts.length === 2) {
          rowMonthYear = String(rawDate).trim();
        }
      }

      if (rowMonthYear !== targetMonth) continue;

      records.push({
        id:           data[i][0],
        date:         _formatDate(data[i][2]),
        checkIn:      data[i][3] ? _formatTime(data[i][3]) : '--',
        checkOut:     data[i][4] ? _formatTime(data[i][4]) : '--',
        remark:       data[i][6] || '',
        status:       data[i][7] || '',
        markedAs:     data[i][8] ? String(data[i][8]) : '',
        workingHours: data[i][9] ? String(data[i][9]) : '--'
      });
    }

    records.sort((a, b) => {
      const [da, ma, ya] = a.date.split('/');
      const [db, mb, yb] = b.date.split('/');
      return new Date(Number(yb), Number(mb)-1, Number(db)) - new Date(Number(ya), Number(ma)-1, Number(da));
    });

    return { success: true, records, month: Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM yyyy') };
  } catch(e) {
    return { success: false, message: e.message, month: '' };
  }
}

// ─── LEAVE MANAGEMENT ────────────────────────────────────────────────────────

function submitLeaveRequest(empCode, empName, leaveType, fromDate, toDate, reason) {
  try {
    // Server-side date validation
    const from = new Date(fromDate);
    const to   = new Date(toDate);
    if (isNaN(from.getTime()) || isNaN(to.getTime())) {
      return { success: false, message: 'Invalid date format.' };
    }
    if (to < from) {
      return { success: false, message: 'End date cannot be before start date.' };
    }
    if (!reason || String(reason).trim().length < 3) {
      return { success: false, message: 'Please provide a reason (min 3 characters).' };
    }
    const validTypes = ['CL', 'SL', 'PL'];
    if (validTypes.indexOf(leaveType) === -1) {
      return { success: false, message: 'Invalid leave type.' };
    }

    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Leave Request");

      // Check balance
      const lbData = _sheetData("Leave Balance");
      let balance = 0;
      const typeMap = { 'CL': 1, 'SL': 2, 'PL': 3 };
      const colIdx  = typeMap[leaveType];

      for (let i = 1; i < lbData.length; i++) {
        if (String(lbData[i][0]) === String(empCode)) {
          balance = Number(lbData[i][colIdx] || 0);
          break;
        }
      }

      // Calculate days
      const days  = Math.round((to - from) / (1000 * 60 * 60 * 24)) + 1;

      if (days <= 0) return { success: false, message: 'Invalid date range.' };
      if (colIdx && balance < days) return { success: false, message: 'Insufficient ' + leaveType + ' balance. Available: ' + balance + ' days.' };

      const leaveId = 'LV' + new Date().getTime();
      const applied = _today();

      sheet.appendRow([leaveId, empCode, empName, leaveType, fromDate, toDate, days, reason, 'Pending', applied, '', '']);

      _writeAuditLog(empCode, 'LEAVE_SUBMITTED', 'LeaveID: ' + leaveId + ', Type: ' + leaveType + ', Days: ' + days, '');
      return { success: true, message: 'Leave request submitted successfully.', leaveId, days };
    } finally {
      _releaseLock(lock);
    }
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getLeaveRequests(empCode, role, filterStatus) {
  try {
    const data = _sheetData("Leave Request");
    const records = [];

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const rowEmp    = String(data[i][1]);
      const rowStatus = String(data[i][8] || '');

      // Employees see only their own; managers/admins see all (or filter by approver)
      if (role === 'Employee' && rowEmp !== String(empCode)) continue;
      if (filterStatus && filterStatus !== 'All' && rowStatus !== filterStatus) continue;

      records.push({
        leaveId:    data[i][0],
        empCode:    data[i][1],
        empName:    data[i][2],
        leaveType:  data[i][3],
        fromDate:   _formatDate(data[i][4]),
        toDate:     _formatDate(data[i][5]),
        days:       data[i][6],
        reason:     data[i][7],
        status:     data[i][8],
        appliedDate:_formatDate(data[i][9]),
        approvedBy: data[i][10] || '',
        approvalDate:_formatDate(data[i][11])
      });
    }
    records.sort((a, b) => b.leaveId.localeCompare(a.leaveId));
    return { success: true, records };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function approveRejectLeave(leaveId, action, approverName) {
  try {
    const lock = _acquireLock();
    try {
      const lrSheet = _getSheet("Leave Request");
      const lrData  = lrSheet.getDataRange().getValues();
      let targetRow = -1, leaveRow = null;

      for (let i = 1; i < lrData.length; i++) {
        if (String(lrData[i][0]) === String(leaveId)) {
          targetRow = i + 1;
          leaveRow  = lrData[i];
          break;
        }
      }
      if (targetRow === -1) return { success: false, message: 'Leave request not found.' };
      if (String(leaveRow[8]) !== 'Pending') return { success: false, message: 'This request is already ' + leaveRow[8] + '.' };

      const newStatus = action === 'Approve' ? 'Approved' : 'Rejected';
      lrSheet.getRange(targetRow, 9).setValue(newStatus);
      lrSheet.getRange(targetRow, 11).setValue(approverName);
      lrSheet.getRange(targetRow, 12).setValue(_today());

      // Deduct balance on approval
      if (newStatus === 'Approved') {
        const typeMap = { 'CL': 2, 'SL': 3, 'PL': 4 };
        const colIdx  = typeMap[String(leaveRow[3])];
        if (colIdx) {
          const lbSheet = _getSheet("Leave Balance");
          const lbData  = lbSheet.getDataRange().getValues();
          for (let i = 1; i < lbData.length; i++) {
            if (String(lbData[i][0]) === String(leaveRow[1])) {
              const current = Number(lbData[i][colIdx - 1] || 0);
              const days    = Number(leaveRow[6] || 0);
              lbSheet.getRange(i + 1, colIdx).setValue(Math.max(0, current - days));
              lbSheet.getRange(i + 1, 5).setValue(_today()); // Last_Updated
              break;
            }
          }
        }
      }

      _writeAuditLog(approverName, 'LEAVE_' + newStatus.toUpperCase(), 'LeaveID: ' + leaveId, '');
      return { success: true, message: 'Leave ' + newStatus.toLowerCase() + ' successfully.' };
    } finally {
      _releaseLock(lock);
    }
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── MASTERS ─────────────────────────────────────────────────────────────────

function getFormMasters() {
  try {
    const result = { departments: [], designations: [], levels: [], processes: [], managers: [] };

    // Department Master
    try {
      const deptData = _sheetData('Department Master');
      for (let i = 1; i < deptData.length; i++) {
        const v = String(deptData[i][0] || '').trim();
        if (v) result.departments.push(v);
      }
    } catch(e) {}

    // Designation Master
    try {
      const desigData = _sheetData('Designation Master');
      for (let i = 1; i < desigData.length; i++) {
        const v = String(desigData[i][0] || '').trim();
        if (v) result.designations.push(v);
      }
    } catch(e) {}

    // Pull unique Levels and Processes from Employee Master
    try {
      const empData = _sheetData("Employee Master");
      const h = empData[0];
      const levelCol   = _colIdx(h, "Level");
      const processCol = _colIdx(h, "Process");
      const levelSet   = new Set();
      const processSet = new Set();
      for (let i = 1; i < empData.length; i++) {
        if (levelCol   >= 0 && empData[i][levelCol])   levelSet.add(String(empData[i][levelCol]).trim());
        if (processCol >= 0 && empData[i][processCol]) processSet.add(String(empData[i][processCol]).trim());
      }
      result.levels    = Array.from(levelSet).sort();
      result.processes = Array.from(processSet).sort();
      if (!result.departments.length) result.departments = result.levels;
      if (!result.designations.length) result.designations = result.levels;
    } catch(e) {}

    // Manager Master
    try {
      const mgrData = _sheetData('Manager Master');
      for (let i = 1; i < mgrData.length; i++) {
        const id   = String(mgrData[i][0] || '').trim();
        const name = String(mgrData[i][1] || '').trim();
        if (id && name) result.managers.push({ id, name });
      }
    } catch(e) {}

    return result;
  } catch(e) {
    return { departments: [], designations: [], levels: [], processes: [], managers: [] };
  }
}

function getHolidays() {
  try {
    const data = _sheetData("Holiday List");
    return data.slice(1).filter(r => r[0]).map(r => ({
      id:   r[0],
      date: _formatDate(r[1]),
      name: r[2],
      type: r[3] || 'Public'
    }));
  } catch(e) {
    return [];
  }
}

function getSalaryDetails(empCode) {
  try {
    const data = _sheetData("Salary Structure");
    if (!data || data.length < 2) return { success: false, message: 'Salary Structure sheet is empty.' };

    const h = data[0];

    // Resolve Emp ID column — try multiple header variants, fall back to col 0
    const empCol = ['Emp_ID','Emp ID','EmpCode','Employee_ID','Emp Code'].reduce((found, name) => {
      return found >= 0 ? found : _colIdx(h, name);
    }, -1);
    const resolvedEmpCol = empCol >= 0 ? empCol : 0;

    // Sheet columns A–K (0-based indices 0–10):
    // A(0): Emp_ID | B(1): Type Of Wages | C(2): Basic | D(3): H.R.A | E(4): Special E.P.F |
    // F(5): CCA | G(6): Conveyance Allowance | H(7): Medical Allowance | I(8): LTA |
    // J(9): Other Allowance | K(10): Gross
    const typeCol  = _colIdx(h, 'Type Of Wages')                                                    >= 0 ? _colIdx(h, 'Type Of Wages')        : 1;
    const basicCol = _colIdx(h, 'Basic')                                                            >= 0 ? _colIdx(h, 'Basic')                : 2;
    const hraCol   = (_colIdx(h, 'H.R.A') >= 0 ? _colIdx(h, 'H.R.A') : _colIdx(h, 'HRA') >= 0 ? _colIdx(h, 'HRA') : 3);
    const sepfCol  = (_colIdx(h, 'Special E.P.F') >= 0 ? _colIdx(h, 'Special E.P.F') :
                      _colIdx(h, 'Special EPF')   >= 0 ? _colIdx(h, 'Special EPF')   : 4);
    const ccaCol   = _colIdx(h, 'CCA')                                                              >= 0 ? _colIdx(h, 'CCA')                  : 5;
    const convCol  = (_colIdx(h, 'Conveyance Allowance') >= 0 ? _colIdx(h, 'Conveyance Allowance') :
                      _colIdx(h, 'Conveyance')            >= 0 ? _colIdx(h, 'Conveyance')           : 6);
    const medCol   = (_colIdx(h, 'Medical Allowance') >= 0 ? _colIdx(h, 'Medical Allowance') :
                      _colIdx(h, 'Medical')            >= 0 ? _colIdx(h, 'Medical')           : 7);
    const ltaCol   = _colIdx(h, 'LTA')                                                              >= 0 ? _colIdx(h, 'LTA')                  : 8;
    const otherCol = (_colIdx(h, 'Other Allowance') >= 0 ? _colIdx(h, 'Other Allowance') :
                      _colIdx(h, 'Other')             >= 0 ? _colIdx(h, 'Other')           : 9);
    const grossCol = (_colIdx(h, 'Gross') >= 0 ? _colIdx(h, 'Gross') :
                      _colIdx(h, 'Gross Salary') >= 0 ? _colIdx(h, 'Gross Salary')         : 10);

    const normCode = String(empCode).trim().toUpperCase();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][resolvedEmpCol] || '').trim().toUpperCase() !== normCode) continue;
      return {
        success:     true,
        empCode:     data[i][resolvedEmpCol],
        typeOfWages: String(data[i][typeCol]  || ''),
        basic:       Number(data[i][basicCol]) || 0,
        hra:         Number(data[i][hraCol])   || 0,
        specialEpf:  Number(data[i][sepfCol])  || 0,
        cca:         Number(data[i][ccaCol])   || 0,
        conveyance:  Number(data[i][convCol])  || 0,
        medical:     Number(data[i][medCol])   || 0,
        lta:         Number(data[i][ltaCol])   || 0,
        other:       Number(data[i][otherCol]) || 0,
        gross:       Number(data[i][grossCol]) || 0
      };
    }

    return {
      success: false,
      message: 'Salary record not found for employee code: ' + empCode +
               '. Please ensure a row exists in the "Salary Structure" sheet with this code in column A.'
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── EMPLOYEE MANAGEMENT ─────────────────────────────────────────────────────

function addNewEmployeeFull(formData) {
  try {
    const lock = _acquireLock();
    try {
      const empSheet   = _getSheet("Employee Master");
      const salSheet   = _getSheet("Salary Structure");
      const loginSheet = _getSheet("Login Details");
      const lbSheet    = _getSheet("Leave Balance");

      // Duplicate check — scan col A for matching Emp ID
      const existingData = empSheet.getDataRange().getValues();
      const alreadyExists = existingData.slice(1).some(r => String(r[0]) === String(formData.empId));
      if (alreadyExists) return { success: false, message: 'Employee ID ' + formData.empId + ' already exists.' };

      // Employee Master columns (matching actual sheet):
      // A: Emp ID | B: Employee Name | C: Employment Role | D: DOB | E: Aadhar Number | F: PAN Number |
      // G: Gender | H: Personal Mobile Number | I: Office Mobile Number | J: Personal Email ID |
      // K: Official Email ID | L: City | M: District | N: State | O: Resign Month | P: Level |
      // Q: Process | R: Account Number | S: IFSC Code | T: Shift ID | U: Status | V: Manager ID |
      // W: Permanent Address | X: Emp Photo URL
      empSheet.appendRow([
        formData.empId,                      // A: Emp ID
        formData.name,                       // B: Employee Name
        formData.empRole       || 'RSPL',    // C: Employment Role
        formData.dob           || '',        // D: DOB
        formData.aadhaar       || '',        // E: Aadhar Number
        formData.pan           || '',        // F: PAN Number
        formData.gender        || '',        // G: Gender
        formData.mobile        || '',        // H: Personal Mobile Number
        formData.offMobile     || '',        // I: Office Mobile Number
        formData.personalEmail || '',        // J: Personal Email ID
        formData.email         || '',        // K: Official Email ID
        formData.city          || '',        // L: City
        formData.district      || '',        // M: District
        formData.state         || '',        // N: State
        formData.resignMonth   || '',        // O: Resign Month
        formData.level         || '',        // P: Level
        formData.process       || '',        // Q: Process
        formData.account       || '',        // R: Account Number
        formData.ifsc          || '',        // S: IFSC Code
        formData.shiftId       || '',        // T: Shift ID
        'Active',                            // U: Status
        formData.managerId     || '',        // V: Manager ID
        formData.address       || '',        // W: Permanent Address
        formData.photoUrl      || ''         // X: Emp Photo URL
      ]);

      salSheet.appendRow([
        formData.empId,
        formData.typeOfWages || '',
        formData.basic       || 0,
        formData.hra         || 0,
        formData.specialEpf  || 0,
        formData.cca         || 0,
        formData.conveyance  || 0,
        formData.medical     || 0,
        formData.lta         || 0,
        formData.other       || 0,
        formData.gross       || 0
      ]);

      // Hash default password
      const hashedPassword = _hashPassword(DEFAULT_PASSWORD);
      loginSheet.appendRow([formData.empId, formData.name, formData.email, hashedPassword]);

      // Initialize leave balance: CL:12, SL:12, PL:15
      lbSheet.appendRow([formData.empId, 12, 12, 15, _today()]);

      _invalidateCache("Employee Master");
      _writeAuditLog(formData.empId, 'EMPLOYEE_ADDED', 'Name: ' + formData.name, '');

      return { success: true, message: 'Employee ' + formData.name + ' added successfully.' };
    } finally {
      _releaseLock(lock);
    }
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getAllEmployees(role) {
  try {
    if (role !== 'Admin' && role !== 'Super Admin' && role !== 'Manager') {
      return { success: false, message: 'Access denied.' };
    }
    const data = _sheetData("Employee Master");
    const h    = data[0];
    const idCol        = _colIdx(h, "Emp ID");
    const nameCol      = _colIdx(h, "Employee Name");
    const empRoleCol   = _colIdx(h, "Employment Role");
    const dobCol       = _colIdx(h, "DOB");
    const aadhaarCol   = _colIdx(h, "Aadhar Number");
    const panCol       = _colIdx(h, "PAN Number");
    const genderCol    = _colIdx(h, "Gender");
    const persMobCol   = _colIdx(h, "Personal Mobile Number");
    const offMobCol    = _colIdx(h, "Office Mobile Number");
    const persEmailCol = _colIdx(h, "Personal Email ID");
    const offEmailCol  = _colIdx(h, "Official Email ID");
    const cityCol      = _colIdx(h, "City");
    const districtCol  = _colIdx(h, "District");
    const stateCol     = _colIdx(h, "State");
    const resignCol    = _colIdx(h, "Resign Month");
    const levelCol     = _colIdx(h, "Level");
    const processCol   = _colIdx(h, "Process");
    const accountCol   = _colIdx(h, "Account Number");
    const ifscCol      = _colIdx(h, "IFSC Code");
    const shiftCol     = _colIdx(h, "Shift ID");
    const statCol      = _colIdx(h, "Status");
    const mgrCol       = _colIdx(h, "Manager ID");
    const addressCol   = _colIdx(h, "Permanent Address");
    const photoCol     = _colIdx(h, "Emp Photo URL") >= 0 ? _colIdx(h, "Emp Photo URL") : 23; // col X fallback

    // Build empCode→name map for manager name lookup
    const empNameMap = {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol]) empNameMap[String(data[i][idCol])] = data[i][nameCol];
    }

    const employees = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][idCol]) continue;
      const mgrId   = mgrCol >= 0 ? String(data[i][mgrCol] || '') : '';
      const mgrName = empNameMap[mgrId] || mgrId || '—';
      employees.push({
        empCode:        String(data[i][idCol]),
        name:           data[i][nameCol],
        employmentRole: empRoleCol   >= 0 ? String(data[i][empRoleCol]   || '') : '',
        dob:            _formatDate(dobCol >= 0 ? data[i][dobCol] : ''),
        aadhaar:        aadhaarCol   >= 0 ? String(data[i][aadhaarCol]   || '') : '',
        pan:            panCol       >= 0 ? String(data[i][panCol]       || '') : '',
        gender:         genderCol    >= 0 ? String(data[i][genderCol]    || '') : '',
        personalMobile: persMobCol   >= 0 ? String(data[i][persMobCol]   || '') : '',
        officialMobile: offMobCol    >= 0 ? String(data[i][offMobCol]    || '') : '',
        personalEmail:  persEmailCol >= 0 ? String(data[i][persEmailCol] || '') : '',
        officialEmail:  offEmailCol  >= 0 ? String(data[i][offEmailCol]  || '') : '',
        city:           cityCol      >= 0 ? String(data[i][cityCol]      || '') : '',
        district:       districtCol  >= 0 ? String(data[i][districtCol]  || '') : '',
        state:          stateCol     >= 0 ? String(data[i][stateCol]     || '') : '',
        resignMonth:    resignCol    >= 0 ? String(data[i][resignCol]    || '') : '',
        level:          levelCol     >= 0 ? String(data[i][levelCol]     || '') : '',
        process:        processCol   >= 0 ? String(data[i][processCol]   || '') : '',
        dept:           levelCol     >= 0 ? String(data[i][levelCol]     || '') : '', // dept mapped to level
        designation:    levelCol     >= 0 ? String(data[i][levelCol]     || '') : '', // designation mapped to level
        accountNumber:  accountCol   >= 0 ? String(data[i][accountCol]   || '') : '',
        ifscCode:       ifscCol      >= 0 ? String(data[i][ifscCol]      || '') : '',
        shiftId:        shiftCol     >= 0 ? String(data[i][shiftCol]     || '') : '',
        status:         statCol      >= 0 ? (String(data[i][statCol] || '') || 'Active') : 'Active',
        managerId:      mgrId,
        managerName:    mgrName,
        address:        addressCol   >= 0 ? String(data[i][addressCol]   || '') : '',
        photoUrl:       _drivePhotoUrl(String(data[i][photoCol] || ''))
      });
    }
    return { success: true, employees };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateEmployeeStatus(empCode, newStatus, updaterRole) {
  try {
    if (updaterRole !== 'Admin' && updaterRole !== 'Super Admin') {
      return { success: false, message: 'Access denied.' };
    }
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Employee Master");
      const data  = sheet.getDataRange().getValues();
      const h     = data[0];
      const statCol = _colIdx(h, "Status");
      if (statCol === -1) return { success: false, message: 'Status column not found.' };
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(empCode)) {
          sheet.getRange(i + 1, statCol + 1).setValue(newStatus);
          _invalidateCache("Employee Master");
          _writeAuditLog(updaterRole, 'STATUS_CHANGED', 'EmpCode: ' + empCode + ' → ' + newStatus, '');
          return { success: true, message: 'Status updated to ' + newStatus };
        }
      }
      return { success: false, message: 'Employee not found.' };
    } finally {
      _releaseLock(lock);
    }
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── EMPLOYEE PHOTO UPDATE ────────────────────────────────────────────────

function updateEmployeePhoto(empCode, photoDataUrl) {
  try {
    const sheet = _getSheet("Employee Master");
    const data  = sheet.getDataRange().getValues();
    const h     = data[0];
    const photoCol = _colIdx(h, "Emp Photo URL"); // 0-based
    const resolvedPhotoCol = photoCol >= 0 ? photoCol : 23; // fallback to col X (index 23)
    if (resolvedPhotoCol === -1) return { success: false, message: 'Emp Photo URL column not found in Employee Master.' };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(empCode)) {
        sheet.getRange(i + 1, resolvedPhotoCol + 1).setValue(photoDataUrl);
        return { success: true, message: 'Photo updated successfully.' };
      }
    }
    return { success: false, message: 'Employee not found.' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── GATE PASS ───────────────────────────────────────────────────────────────
// Sheet: "Gate Pass" — columns:
// PassID | EmpCode | EmpName | PassType | Date | TimeOut | TimeIn | Reason | Status | AppliedDate | ApprovedBy | ApprovalDate

function submitGatePass(empCode, empName, passType, date, timeOut, timeIn, reason) {
  try {
    const sheet  = _getSheet("Gate Pass");
    const passId = 'GP' + new Date().getTime();
    sheet.appendRow([
      passId, empCode, empName, passType,
      date, timeOut, timeIn || '', reason,
      'Pending', _today(), '', ''
    ]);
    return { success: true, message: 'Gate pass request submitted successfully.' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getGatePasses(empCode, role, filterStatus) {
  try {
    const data = _sheetData("Gate Pass");
    if (data.length <= 1) return { success: true, records: [] }; // empty sheet

    const records = [];

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      const rowEmp    = String(data[i][1] || '').trim();
      const rowStatus = String(data[i][8] || '').trim();

      // Employees see only their own; admins/managers see all
      if (role === 'Employee' && rowEmp !== String(empCode).trim()) continue;
      if (filterStatus && filterStatus !== 'All' && rowStatus !== filterStatus) continue;

      // Date stored as YYYY-MM-DD string from HTML date input — normalise for display
      const rawDate = data[i][4];
      let displayDate = '';
      if (rawDate instanceof Date) {
        displayDate = _formatDate(rawDate);
      } else {
        const s = String(rawDate || '').trim();
        // Convert YYYY-MM-DD → DD/MM/YYYY
        if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
          const [y, m, d] = s.split('-');
          displayDate = d + '/' + m + '/' + y;
        } else {
          displayDate = s;
        }
      }

      records.push({
        passId:      String(data[i][0]),
        empCode:     String(data[i][1] || ''),
        empName:     String(data[i][2] || ''),
        passType:    String(data[i][3] || ''),
        date:        displayDate,
        timeOut:     data[i][5] ? _formatTime(data[i][5]) : '—',
        timeIn:      data[i][6] ? _formatTime(data[i][6]) : '—',
        reason:      String(data[i][7] || ''),
        status:      rowStatus || 'Pending',
        appliedDate: _formatDate(data[i][9]),
        approvedBy:  String(data[i][10] || ''),
        approvalDate:_formatDate(data[i][11])
      });
    }

    records.sort((a, b) => b.passId.localeCompare(a.passId));
    return { success: true, records };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function approveRejectGatePass(passId, action, approverName) {
  try {
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Gate Pass");
      const data  = sheet.getDataRange().getValues();
      let targetRow = -1, rowData = null;

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(passId)) {
          targetRow = i + 1;
          rowData   = data[i];
          break;
        }
      }
      if (targetRow === -1) return { success: false, message: 'Gate pass not found.' };
      if (String(rowData[8]) !== 'Pending') return { success: false, message: 'Already ' + rowData[8] + '.' };

      const newStatus = action === 'Approve' ? 'Approved' : 'Rejected';
      sheet.getRange(targetRow, 9).setValue(newStatus);
      sheet.getRange(targetRow, 11).setValue(approverName);
      sheet.getRange(targetRow, 12).setValue(_today());

      _writeAuditLog(approverName, 'GATEPASS_' + newStatus.toUpperCase(), 'PassID: ' + passId, '');
      return { success: true, message: 'Gate pass ' + newStatus.toLowerCase() + ' successfully.' };
    } finally {
      _releaseLock(lock);
    }
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── POLICIES ────────────────────────────────────────────────────────────────
// Sheet: "Policies" — columns:
// PolicyID | Title | Category | EffectiveDate | Version | Description | Tags | DriveFileId | FileName | UploadedBy | UploadedDate | MimeType
//
// Files are stored in Google Drive under "HRMS_Policies/" folder.
// Max upload size: 30 MB (enforced on the client side).

/** Get or create the root HRMS Policies folder in Drive */
function _getPoliciesDriveFolder() {
  const FOLDER_NAME = 'HRMS_Policies';
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(FOLDER_NAME);
}

function uploadPolicy(data) {
  try {
    if (data.uploaderRole !== 'Admin' && data.uploaderRole !== 'Super Admin' && data.uploaderRole !== 'Manager') {
      return { success: false, message: 'Access denied.' };
    }

    const sheet    = _getSheet('Policies');
    const policyId = 'POL' + new Date().getTime();
    const fileName = data.fileName || ('policy_' + policyId);
    const mimeType = data.mimeType || 'application/octet-stream';

    // ── Store file in Google Drive (supports up to 30 MB) ──────
    let driveFileId = '';
    const base64Data = data.fileData || '';
    if (base64Data) {
      const blob      = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
      const folder    = _getPoliciesDriveFolder();
      const driveFile = folder.createFile(blob);
      driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      driveFileId = driveFile.getId();
    }

    // Columns: PolicyID | Title | Category | EffectiveDate | Version | Description | Tags | DriveFileId | FileName | UploadedBy | UploadedDate | MimeType
    sheet.appendRow([
      policyId,
      data.title,
      data.category,
      data.effectiveDate || '',
      data.version       || '',
      data.description   || '',
      data.tags          || '',
      driveFileId,
      fileName,
      data.uploadedBy,
      _today(),
      mimeType
    ]);

    return { success: true, message: 'Policy uploaded successfully.' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getPolicies(category) {
  try {
    const data     = _sheetData("Policies");
    const policies = [];

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const rowCat = String(data[i][2] || '');
      if (category !== 'All' && rowCat !== category) continue;

      // col 7 = DriveFileId, col 8 = fileName
      // Support legacy rows that stored raw base64 in col 7
      const col7     = String(data[i][7] || '');
      const fileName = String(data[i][8] || '');
      let fileData   = '';

      if (col7.startsWith('data:') || col7.length > 100) {
        // Legacy: raw base64 data URL stored directly in sheet
        fileData = col7;
      } else if (col7) {
        // New: Drive file ID — build a direct view URL
        fileData = 'https://drive.google.com/uc?export=view&id=' + col7;
      }

      policies.push({
        policyId:      String(data[i][0]),
        title:         String(data[i][1] || ''),
        category:      String(data[i][2] || ''),
        effectiveDate: _formatDate(data[i][3]),
        version:       String(data[i][4] || ''),
        description:   String(data[i][5] || ''),
        tags:          String(data[i][6] || ''),
        fileData:      fileData,   // Drive URL or legacy base64 — used by frontend to open/download
        fileName:      fileName,
        uploadedBy:    String(data[i][9] || ''),
        uploadedDate:  _formatDate(data[i][10])
      });
    }

    policies.sort((a, b) => b.policyId.localeCompare(a.policyId));
    return { success: true, policies };
  } catch(e) {
    return { success: false, message: e.message, policies: [] };
  }
}

function deletePolicy(policyId, role) {
  try {
    if (role !== 'Admin' && role !== 'Super Admin' && role !== 'Manager') {
      return { success: false, message: 'Access denied.' };
    }
    const sheet = _getSheet("Policies");
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(policyId)) {
        // Delete the Drive file if one exists (col 7 = DriveFileId)
        const driveFileId = String(data[i][7] || '');
        if (driveFileId && !driveFileId.startsWith('data:') && driveFileId.length < 100) {
          try { DriveApp.getFileById(driveFileId).setTrashed(true); } catch(e) {}
        }
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Policy deleted.' };
      }
    }
    return { success: false, message: 'Policy not found.' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── UPDATE OFFICIAL DETAILS (Manager / Admin / Super Admin only) ─────────────

function updateOfficialDetails(empCode, fields, updaterRole) {
  try {
    if (updaterRole !== 'Manager' && updaterRole !== 'Admin' && updaterRole !== 'Super Admin') {
      return { success: false, message: 'Access denied.' };
    }
    const sheet = _getSheet("Employee Master");
    const data  = sheet.getDataRange().getValues();
    const h     = data[0];
    // Use flexible column matching (case-insensitive, ignores underscores/spaces)
    const offEmailCol  = _colIdx(h, "Official Email ID");
    const offMobCol    = _colIdx(h, "Office Mobile Number");
    const levelCol     = _colIdx(h, "Level");
    const processCol   = _colIdx(h, "Process");
    const empRoleCol   = _colIdx(h, "Employment Role");
    const shiftCol     = _colIdx(h, "Shift ID");
    const mgrCol       = _colIdx(h, "Manager ID");

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(empCode)) {
        const row = i + 1;
        if (fields.officeEmail    !== undefined && offEmailCol >= 0) sheet.getRange(row, offEmailCol + 1).setValue(fields.officeEmail);
        if (fields.officeMobile   !== undefined && offMobCol   >= 0) sheet.getRange(row, offMobCol   + 1).setValue(fields.officeMobile);
        if (fields.level          !== undefined && levelCol    >= 0) sheet.getRange(row, levelCol    + 1).setValue(fields.level);
        if (fields.process        !== undefined && processCol  >= 0) sheet.getRange(row, processCol  + 1).setValue(fields.process);
        if (fields.department     !== undefined && levelCol    >= 0) sheet.getRange(row, levelCol    + 1).setValue(fields.department); // legacy compat
        if (fields.designation    !== undefined && levelCol    >= 0) sheet.getRange(row, levelCol    + 1).setValue(fields.designation); // legacy compat
        if (fields.employmentRole !== undefined && empRoleCol  >= 0) sheet.getRange(row, empRoleCol  + 1).setValue(fields.employmentRole);
        if (fields.shiftId        !== undefined && shiftCol    >= 0) sheet.getRange(row, shiftCol    + 1).setValue(fields.shiftId);
        if (fields.managerId      !== undefined && mgrCol      >= 0) sheet.getRange(row, mgrCol      + 1).setValue(fields.managerId);
        _invalidateCache("Employee Master");
        return { success: true, message: 'Official details updated successfully.' };
      }
    }
    return { success: false, message: 'Employee not found.' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── UPDATE SALARY DETAILS (Manager / Admin / Super Admin only) ───────────────

function updateSalaryDetails(empCode, salaryData, updaterRole) {
  try {
    if (updaterRole !== 'Manager' && updaterRole !== 'Admin' && updaterRole !== 'Super Admin') {
      return { success: false, message: 'Access denied.' };
    }
    const sheet = _getSheet("Salary Structure");
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(empCode)) {
        // Write cols B–K (1-based 2–11) — 10 columns
        // B: Type Of Wages | C: Basic | D: H.R.A | E: Special E.P.F |
        // F: CCA | G: Conveyance Allowance | H: Medical Allowance | I: LTA |
        // J: Other Allowance | K: Gross
        sheet.getRange(i + 1, 2, 1, 10).setValues([[
          String(salaryData.typeOfWages  || ''),
          Number(salaryData.basic)       || 0,
          Number(salaryData.hra)         || 0,
          Number(salaryData.specialEpf)  || 0,
          Number(salaryData.cca)         || 0,
          Number(salaryData.conveyance)  || 0,
          Number(salaryData.medical)     || 0,
          Number(salaryData.lta)         || 0,
          Number(salaryData.other)       || 0,
          Number(salaryData.gross)       || 0
        ]]);
        return { success: true, message: 'Salary updated successfully.' };
      }
    }
    return { success: false, message: 'Salary record not found.' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── TASKS ────────────────────────────────────────────────────────────────────
// Sheet: "Task Sheet" — columns (0-based):
// 0: Doc Id | 1: Blog Name | 2: Deadline | 3: RMS Name | 4: Doer Name |
// 5: Email ID | 6: Action Link | 7: Dashboard (CLICK HERE link)
// Employees see tasks where Email ID or Doer Name matches their record.

function getMyTasks(empCode, empEmail) {
  try {
    const data = _sheetData("Task Sheet");
    if (data.length <= 1) return { success: true, tasks: [] };

    const h = data[0];
    // Resolve column index by header name (case-insensitive)
    const col = name => {
      const n = name.toLowerCase().trim();
      const idx = h.findIndex(c => String(c).toLowerCase().trim() === n);
      return idx;
    };

    const docIdCol    = col('doc id')       >= 0 ? col('doc id')       : 0;
    const blogCol     = col('blog name')    >= 0 ? col('blog name')    : 1;
    const deadlineCol = col('deadline')     >= 0 ? col('deadline')     : 2;
    const rmsCol      = col('rms name')     >= 0 ? col('rms name')     : 3;
    const doerCol     = col('doer name')    >= 0 ? col('doer name')    : 4;
    const emailCol    = col('email id')     >= 0 ? col('email id')     : 5;
    const actionCol   = col('action link')  >= 0 ? col('action link')  : 6;
    const dashCol     = col('dashboard')    >= 0 ? col('dashboard')    : 7;
    const statusCol   = col('status')       >= 0 ? col('status')       : -1;

    // Look up employee's full name from Employee Master for name-based matching
    let empName = '';
    try {
      const empData = _sheetData("Employee Master");
      const empH    = empData[0];
      const idIdx   = _colIdx(empH, 'Emp ID');
      const nmIdx   = _colIdx(empH, 'Employee Name');
      for (let i = 1; i < empData.length; i++) {
        if (String(empData[i][idIdx]).trim() === String(empCode).trim()) {
          empName = String(empData[i][nmIdx] || '').trim().toLowerCase();
          break;
        }
      }
    } catch(e) {}

    const empCodeStr = String(empCode || '').trim().toLowerCase();
    const tasks = [];

    for (let i = 1; i < data.length; i++) {
      if (!data[i][docIdCol] && !data[i][blogCol]) continue;

      const rowDoer = String(data[i][doerCol] || '').trim();

      // Doer Name format: "Pratul Sharma - RS367"  (Name - EmpCode)
      // Extract emp code (after last " - ") and name (before last " - ")
      const dashIdx   = rowDoer.lastIndexOf(' - ');
      const doerEmpId = dashIdx >= 0
        ? rowDoer.substring(dashIdx + 3).trim().toLowerCase()
        : '';
      const doerName  = dashIdx >= 0
        ? rowDoer.substring(0, dashIdx).trim().toLowerCase()
        : rowDoer.toLowerCase();

      // Match ONLY by emp code or emp name — no email matching
      const matched =
        (empCodeStr && doerEmpId && doerEmpId === empCodeStr) ||
        (empName    && doerName  && doerName  === empName);

      if (!matched) continue;

      const rawDeadline = data[i][deadlineCol];
      let deadlineStr = '';
      if (rawDeadline instanceof Date) {
        deadlineStr = _formatDate(rawDeadline) +
          (rawDeadline.getHours() || rawDeadline.getMinutes()
            ? ' ' + _formatTime(rawDeadline) : '');
      } else {
        deadlineStr = String(rawDeadline || '').trim();
      }

      tasks.push({
        docId:         String(data[i][docIdCol]  || ''),
        blogName:      String(data[i][blogCol]   || ''),
        deadline:      deadlineStr,
        rmsName:       String(data[i][rmsCol]    || ''),
        doerName:      String(data[i][doerCol]   || ''),
        email:         String(data[i][emailCol]  || ''),
        actionLink:    String(data[i][actionCol] || ''),
        dashboardLink: dashCol >= 0 ? String(data[i][dashCol] || '') : '',
        status:        statusCol >= 0 ? String(data[i][statusCol] || '') : ''
      });
    }

    return { success: true, tasks };
  } catch(e) {
    return { success: false, message: e.message, tasks: [] };
  }
}

// ─── CHANGE PASSWORD ─────────────────────────────────────────────────────────

function changePassword(empCode, currentPassword, newPassword) {
  try {
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Login Details");
      const data  = sheet.getDataRange().getValues();
      const hashedCurrent = _hashPassword(currentPassword);
      const hashedNew = _hashPassword(newPassword);

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(empCode).trim()) {
          // Verify current password (supports both plain-text and hashed)
          const storedPwd = String(data[i][3]);
          const isMatch = storedPwd === currentPassword || storedPwd === hashedCurrent;
          if (!isMatch) {
            return { success: false, message: 'Current password is incorrect.' };
          }
          // Update to new hashed password
          sheet.getRange(i + 1, 4).setValue(hashedNew);
          _writeAuditLog(empCode, 'PASSWORD_CHANGED', 'Password updated', '');
          return { success: true, message: 'Password changed successfully.' };
        }
      }
      return { success: false, message: 'Employee record not found.' };
    } finally {
      _releaseLock(lock);
    }
  } catch(e) {
    return { success: false, message: e.message };
  }
}


// ─── DOCUMENT MANAGEMENT (Google Drive backend — supports up to 10 MB) ────────
//
// Sheet: "Employee Documents" — columns (0-based):
// 0: DocID | 1: EmpCode | 2: EmpName | 3: DocType | 4: FileName |
// 5: DriveFileId | 6: UploadedBy | 7: UploadedDate | 8: Notes | 9: MimeType
//
// Files are stored in Google Drive under a folder "HRMS_Employee_Documents/<EmpCode>/".
// The sheet stores only the Drive file ID — no base64 in cells.
// Max upload size: 10 MB (enforced on the client side).

/** Get or create the root HRMS documents folder in Drive */
function _getHrmsDriveFolder() {
  const ROOT_FOLDER_NAME = 'HRMS_Employee_Documents';
  const folders = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(ROOT_FOLDER_NAME);
}

/** Get or create a per-employee subfolder */
function _getEmpDriveFolder(empCode) {
  const root = _getHrmsDriveFolder();
  const sub  = root.getFoldersByName(empCode);
  if (sub.hasNext()) return sub.next();
  return root.createFolder(empCode);
}

/** Check manager-team membership (shared helper) */
function _isTeamMember(managerCode, empCode) {
  const empData = _sheetData("Employee Master");
  const h = empData[0];
  const idCol  = _colIdx(h, "Emp ID");
  const mgrCol = _colIdx(h, "Manager ID");
  for (let i = 1; i < empData.length; i++) {
    if (String(empData[i][idCol]) === empCode) {
      return String(empData[i][mgrCol]) === managerCode;
    }
  }
  return false;
}

/**
 * Upload a document for an employee — stores file in Google Drive.
 * docData: { empCode, empName, docType, fileName, mimeType,
 *            fileDataBase64 (pure base64, no data: prefix),
 *            uploadedBy, uploaderCode, uploaderRole, notes }
 */
function uploadEmployeeDocument(docData) {
  try {
    const uploaderRole = docData.uploaderRole || 'Employee';
    const uploaderCode = String(docData.uploaderCode || '').trim();
    const targetEmp    = String(docData.empCode || '').trim();

    // ── Access control ──────────────────────────────────────────
    if (uploaderRole === 'Employee' && uploaderCode !== targetEmp) {
      return { success: false, message: 'Access denied. You can only upload your own documents.' };
    }
    if (uploaderRole === 'Manager' && uploaderCode !== targetEmp) {
      if (!_isTeamMember(uploaderCode, targetEmp)) {
        return { success: false, message: 'Access denied. Employee is not in your team.' };
      }
    }

    // ── Decode base64 and save to Drive ─────────────────────────
    const base64Data = docData.fileDataBase64 || '';
    if (!base64Data) return { success: false, message: 'No file data received.' };

    const mimeType = docData.mimeType || 'application/octet-stream';
    const fileName = docData.fileName || ('document_' + new Date().getTime());

    const blob    = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    const folder  = _getEmpDriveFolder(targetEmp);
    const driveFile = folder.createFile(blob);
    // Share with anyone who has the link — required for mobile access without Google sign-in
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    // Also ensure the file is not restricted by the parent folder's domain policy
    try { driveFile.setShareableByEditors(false); } catch(e) {}

    const driveFileId = driveFile.getId();

    // ── Resolve employee name ────────────────────────────────────
    let empName = docData.empName || '';
    if (!empName) {
      try {
        const empData = _sheetData("Employee Master");
        const h = empData[0];
        const idCol   = _colIdx(h, "Emp ID");
        const nameCol = _colIdx(h, "Employee Name");
        for (let i = 1; i < empData.length; i++) {
          if (String(empData[i][idCol]) === targetEmp) {
            empName = String(empData[i][nameCol] || '');
            break;
          }
        }
      } catch(e) {}
    }

    // ── Write metadata row to sheet ──────────────────────────────
    const sheet = _getSheet("Employee Documents");
    const docId = 'DOC' + new Date().getTime();

    sheet.appendRow([
      docId,
      targetEmp,
      empName,
      docData.docType  || 'Other',
      fileName,
      driveFileId,
      docData.uploadedBy || uploaderCode,
      _today(),
      docData.notes    || '',
      mimeType
    ]);

    return { success: true, message: 'Document uploaded successfully.', docId, driveFileId };
  } catch(e) {
    return { success: false, message: 'Upload failed: ' + e.message };
  }
}

/**
 * Get all documents for a given employee.
 * Returns metadata + a short-lived Drive view URL (no base64 in response).
 */
function getEmployeeDocuments(empCode, requestorCode, requestorRole) {
  try {
    const target = String(empCode || '').trim();

    // ── Access control ──────────────────────────────────────────
    if (requestorRole === 'Employee' && requestorCode !== target) {
      return { success: false, message: 'Access denied.' };
    }
    if (requestorRole === 'Manager' && requestorCode !== target) {
      if (!_isTeamMember(requestorCode, target)) {
        return { success: false, message: 'Access denied. Employee is not in your team.' };
      }
    }

    const data = _sheetData("Employee Documents");
    if (data.length <= 1) return { success: true, documents: [] };

    const documents = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (String(data[i][1]).trim() !== target) continue;

      const driveFileId = String(data[i][5] || '');

      // Build view/download URLs from Drive file ID
      // Using uc?export=view avoids the "permission" prompt on mobile browsers
      let viewUrl     = '';
      let downloadUrl = '';
      if (driveFileId) {
        viewUrl     = 'https://drive.google.com/uc?export=view&id=' + driveFileId;
        downloadUrl = 'https://drive.google.com/uc?export=download&id=' + driveFileId;
      }

      documents.push({
        docId:        String(data[i][0]),
        empCode:      String(data[i][1] || ''),
        empName:      String(data[i][2] || ''),
        docType:      String(data[i][3] || 'Other'),
        fileName:     String(data[i][4] || ''),
        driveFileId:  driveFileId,
        viewUrl:      viewUrl,
        downloadUrl:  downloadUrl,
        uploadedBy:   String(data[i][6] || ''),
        uploadedDate: _formatDate(data[i][7]),
        notes:        String(data[i][8] || ''),
        mimeType:     String(data[i][9] || '')
      });
    }

    documents.sort((a, b) => b.docId.localeCompare(a.docId));
    return { success: true, documents };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Fetch a Drive file and return it as a base64 data URL.
 * This proxies the file through Apps Script so mobile users never hit
 * a Google Drive permission/sign-in wall.
 *
 * Returns: { success: true, dataUrl: 'data:<mime>;base64,...', mimeType, fileName }
 *       or { success: false, message: '...' }
 */
function getDocumentAsBase64(docId, requestorCode, requestorRole) {
  try {
    const target = String(docId || '').trim();

    // Look up the doc record to get driveFileId and access-check the owner
    const data = _sheetData("Employee Documents");
    let driveFileId = '';
    let mimeType    = 'application/octet-stream';
    let fileName    = 'document';
    let docOwner    = '';

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== target) continue;
      docOwner    = String(data[i][1]).trim();
      driveFileId = String(data[i][5] || '').trim();
      fileName    = String(data[i][4] || 'document');
      mimeType    = String(data[i][9] || 'application/octet-stream');
      break;
    }

    if (!driveFileId) return { success: false, message: 'Document not found.' };

    // Access control
    if (requestorRole === 'Employee' && requestorCode !== docOwner) {
      return { success: false, message: 'Access denied.' };
    }
    if (requestorRole === 'Manager' && requestorCode !== docOwner) {
      if (!_isTeamMember(requestorCode, docOwner)) {
        return { success: false, message: 'Access denied.' };
      }
    }

    // Fetch from Drive and encode as base64
    const file    = DriveApp.getFileById(driveFileId);
    const blob    = file.getBlob();
    const base64  = Utilities.base64Encode(blob.getBytes());
    const resolvedMime = blob.getContentType() || mimeType;
    const dataUrl = 'data:' + resolvedMime + ';base64,' + base64;

    return { success: true, dataUrl: dataUrl, mimeType: resolvedMime, fileName: fileName };
  } catch(e) {
    return { success: false, message: 'Could not load document: ' + e.message };
  }
}

/**
 * Fetch a Drive photo by its Drive file URL and return it as a base64 data URL.
 * Used by the frontend to display employee photos without hitting Drive auth walls on mobile.
 * driveUrl: any Google Drive sharing URL (file/d/ID/view, uc?id=, thumbnail?id=, etc.)
 */
function getPhotoAsBase64(driveUrl) {
  try {
    if (!driveUrl) return { success: false };
    const s = String(driveUrl).trim();

    // Extract file ID from any Drive URL format
    let fileId = '';
    const m1 = s.match(/drive\.google\.com\/file\/d\/([^\/\?]+)/);
    if (m1) fileId = m1[1];
    if (!fileId) { const m2 = s.match(/[?&]id=([^&]+)/); if (m2) fileId = m2[1]; }
    if (!fileId) return { success: false, message: 'Not a Drive URL' };

    const file   = DriveApp.getFileById(fileId);
    const blob   = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    const mime   = blob.getContentType() || 'image/jpeg';
    return { success: true, dataUrl: 'data:' + mime + ';base64,' + base64 };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Delete a document — removes from Drive AND from the sheet row.
 */
function deleteEmployeeDocument(docId, requestorCode, requestorRole) {
  try {
    const sheet = _getSheet("Employee Documents");
    const data  = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(docId)) continue;

      const docOwner    = String(data[i][1]).trim();
      const driveFileId = String(data[i][5] || '');

      // ── Access control ────────────────────────────────────────
      if (requestorRole === 'Employee' && requestorCode !== docOwner) {
        return { success: false, message: 'Access denied.' };
      }
      if (requestorRole === 'Manager' && requestorCode !== docOwner) {
        if (!_isTeamMember(requestorCode, docOwner)) {
          return { success: false, message: 'Access denied.' };
        }
      }

      // ── Delete from Drive ─────────────────────────────────────
      if (driveFileId) {
        try { DriveApp.getFileById(driveFileId).setTrashed(true); } catch(e) {}
      }

      // ── Delete sheet row ──────────────────────────────────────
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Document deleted.' };
    }
    return { success: false, message: 'Document not found.' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── MANAGER DASHBOARD ───────────────────────────────────────────────────────

/**
 * Get all employees under a given manager.
 * Returns employee list with basic stats.
 */
function getManagerTeam(managerCode) {
  try {
    const empData = _sheetData("Employee Master");
    const h = empData[0];
    const idCol      = _colIdx(h, "Emp ID");
    const nameCol    = _colIdx(h, "Employee Name");
    const levelCol   = _colIdx(h, "Level");
    const processCol = _colIdx(h, "Process");
    const statusCol  = _colIdx(h, "Status");
    const mgrCol     = _colIdx(h, "Manager ID");
    const dobCol     = _colIdx(h, "DOB");
    const dojCol     = _colIdx(h, "Date of Joining");
    const photoCol   = _colIdx(h, "Emp Photo URL") >= 0 ? _colIdx(h, "Emp Photo URL") : 23; // col X fallback

    const team = [];
    for (let i = 1; i < empData.length; i++) {
      if (!empData[i][idCol]) continue;
      const rowMgr = String(empData[i][mgrCol] || '').trim();
      if (rowMgr !== String(managerCode).trim()) continue;

      team.push({
        empCode:     String(empData[i][idCol]),
        name:        String(empData[i][nameCol]   || ''),
        dept:        levelCol   >= 0 ? String(empData[i][levelCol]   || '') : '',
        designation: levelCol   >= 0 ? String(empData[i][levelCol]   || '') : '',
        level:       levelCol   >= 0 ? String(empData[i][levelCol]   || '') : '',
        process:     processCol >= 0 ? String(empData[i][processCol] || '') : '',
        status:      statusCol  >= 0 ? (String(empData[i][statusCol] || '') || 'Active') : 'Active',
        joiningDate: _formatDate(dojCol >= 0 ? empData[i][dojCol] : ''),
        photoUrl:    _drivePhotoUrl(String(empData[i][photoCol] || ''))
      });
    }

    // Attendance summary for current month for each team member
    const attData = _sheetData("Attendance Sheet");
    const now = new Date();
    const monthStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MM/yyyy');

    const attMap = {}; // empCode → { present, absent, late }
    for (let i = 1; i < attData.length; i++) {
      const rowEmp = String(attData[i][1] || '').trim();
      if (!attMap[rowEmp]) attMap[rowEmp] = { present: 0, absent: 0, late: 0 };

      let rowMonthYear = '';
      const rawDate = attData[i][2];
      if (rawDate instanceof Date) {
        if (rawDate.getFullYear() >= 1900) {
          rowMonthYear = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'MM/yyyy');
        }
      } else {
        const parts = String(rawDate).trim().split('/');
        if (parts.length === 3) rowMonthYear = parts[1].padStart(2,'0') + '/' + parts[2];
      }
      if (rowMonthYear !== monthStr) continue;

      const st = String(attData[i][7] || '').toLowerCase().trim();
      if (st === 'absent') {
        attMap[rowEmp].absent++;
      } else if (st === 'late') {
        attMap[rowEmp].present++;
        attMap[rowEmp].late++;
      } else if (attData[i][3]) {
        attMap[rowEmp].present++;
      }
    }

    // Merge attendance into team
    team.forEach(emp => {
      const att = attMap[emp.empCode] || { present: 0, absent: 0, late: 0 };
      emp.monthPresent = att.present;
      emp.monthAbsent  = att.absent;
      emp.monthLate    = att.late;
    });

    const totalEmp  = team.length;
    const activeEmp = team.filter(e => e.status.toLowerCase() === 'active').length;

    // Pending leaves for this manager's team
    let pendingLeaves = 0;
    try {
      const lrData = _sheetData("Leave Request");
      const teamCodes = new Set(team.map(e => e.empCode));
      for (let i = 1; i < lrData.length; i++) {
        if (String(lrData[i][8] || '').toLowerCase() === 'pending' &&
            teamCodes.has(String(lrData[i][1]))) {
          pendingLeaves++;
        }
      }
    } catch(e) {}

    return {
      success: true,
      team,
      stats: { totalEmp, activeEmp, pendingLeaves }
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── AUDIT LOG ───────────────────────────────────────────────────────────────
// Sheet: "Audit Log" — columns: Timestamp | EmpCode | Action | Details | IPAddress

function _writeAuditLog(empCode, action, details, ipAddress) {
  try {
    const sheet = _getSheet("Audit Log");
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    sheet.appendRow([timestamp, empCode, action, details, ipAddress || 'N/A']);
  } catch(e) {
    // Fail silently — audit log is non-critical
  }
}

function getAuditLog(requestorRole, filterEmpCode, filterAction, limit) {
  try {
    if (requestorRole !== 'Admin' && requestorRole !== 'Super Admin') {
      return { success: false, message: 'Access denied.' };
    }
    const data = _sheetData("Audit Log");
    const logs = [];
    const maxLimit = limit || 100;
    
    for (let i = data.length - 1; i >= 1 && logs.length < maxLimit; i--) {
      if (!data[i][0]) continue;
      const rowEmp = String(data[i][1] || '');
      const rowAction = String(data[i][2] || '');
      
      if (filterEmpCode && rowEmp !== filterEmpCode) continue;
      if (filterAction && filterAction !== 'All' && rowAction !== filterAction) continue;
      
      logs.push({
        timestamp: String(data[i][0]),
        empCode: rowEmp,
        action: rowAction,
        details: String(data[i][3] || ''),
        ipAddress: String(data[i][4] || 'N/A')
      });
    }
    
    return { success: true, logs };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function clearAuditLog(requestorRole) {
  if (requestorRole !== 'Super Admin') {
    return { success: false, message: 'Access denied. Only Super Admin can clear the audit log.' };
  }
  try {
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Audit Log");
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return { success: true, message: 'Audit log is already empty.' };
      // Delete all data rows, keep header row
      sheet.deleteRows(2, lastRow - 1);
      // Write a record of the clear action itself
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
      sheet.appendRow([timestamp, 'SYSTEM', 'AUDIT_LOG_CLEARED', 'All previous logs cleared by Super Admin', 'N/A']);
      return { success: true, message: 'Audit log cleared successfully.' };
    } finally {
      _releaseLock(lock);
    }
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── EMAIL NOTIFICATIONS (removed) ──────────────────────────────────────────

// ─── ANNOUNCEMENTS ───────────────────────────────────────────────────────────
// Sheet: "Announcements" — columns: AnnID | Title | Content | Category | PostedBy | PostedDate | ExpiryDate | Priority

function postAnnouncement(data, posterRole) {
  try {
    if (posterRole !== 'Admin' && posterRole !== 'Super Admin' && posterRole !== 'Manager') {
      return { success: false, message: 'Access denied.' };
    }
    const lock = _acquireLock();
    try {
      // Auto-create the Announcements sheet if it doesn't exist
      let sheet = _getSpreadsheet().getSheetByName("Announcements");
      if (!sheet) {
        sheet = _getSpreadsheet().insertSheet("Announcements");
        sheet.appendRow(['AnnID', 'Title', 'Content', 'Category', 'PostedBy', 'PostedDate', 'ExpiryDate', 'Priority']);
      }
      const annId = 'ANN' + new Date().getTime();
      // Format expiryDate from ISO (yyyy-MM-dd) to dd/MM/yyyy if provided
      let expiryFormatted = '';
      if (data.expiryDate) {
        const parts = String(data.expiryDate).split('-');
        if (parts.length === 3) {
          expiryFormatted = parts[2] + '/' + parts[1] + '/' + parts[0]; // dd/MM/yyyy
        } else {
          expiryFormatted = data.expiryDate;
        }
      }

      sheet.appendRow([
        annId,
        data.title || '',
        data.content || '',
        data.category || 'General',
        data.postedBy || '',          // stores the poster's name
        _today(),
        expiryFormatted,
        data.priority || 'Normal'
      ]);
      _invalidateCache("Announcements");
      _writeAuditLog(data.postedByCode || data.postedBy, 'POST_ANNOUNCEMENT', 'Posted: ' + data.title, '');
      return { success: true, message: 'Announcement posted successfully.', annId };
    } finally {
      _releaseLock(lock);
    }
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getAnnouncements(category) {
  try {
    // Gracefully handle missing sheet — return empty list instead of error
    const annSheet = _getSpreadsheet().getSheetByName("Announcements");
    if (!annSheet) return { success: true, announcements: [] };

    const data = annSheet.getDataRange().getValues();
    const announcements = [];
    const today = new Date();
    
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const rowCat = String(data[i][3] || '');
      if (category !== 'All' && rowCat !== category) continue;
      
      // Check expiry — handles both Date objects and dd/MM/yyyy strings
      const expiry = data[i][6];
      if (expiry) {
        let expiryDate;
        if (expiry instanceof Date) {
          // Guard against epoch-zero dates (empty cells sometimes read as Date(0))
          if (expiry.getFullYear() >= 1970) {
            expiryDate = expiry;
          }
        } else {
          const s = String(expiry).trim();
          if (s) {
            // dd/MM/yyyy → parse manually
            const parts = s.split('/');
            if (parts.length === 3) {
              expiryDate = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
            } else {
              // ISO format yyyy-MM-dd or similar
              const isoParts = s.split('-');
              if (isoParts.length === 3) {
                expiryDate = new Date(Number(isoParts[0]), Number(isoParts[1]) - 1, Number(isoParts[2]));
              } else {
                expiryDate = new Date(s);
              }
            }
          }
        }
        // Set expiry to end of that day (23:59:59) so it shows all day on expiry date
        if (expiryDate && !isNaN(expiryDate.getTime())) {
          expiryDate.setHours(23, 59, 59, 999);
          if (expiryDate < today) continue; // skip only if fully past
        }
      }
      
      announcements.push({
        annId: String(data[i][0]),
        title: String(data[i][1] || ''),
        content: String(data[i][2] || ''),
        category: String(data[i][3] || 'General'),
        postedBy: String(data[i][4] || ''),
        postedDate: _formatDate(data[i][5]),
        expiryDate: _formatDate(data[i][6]),
        priority: String(data[i][7] || 'Normal')
      });
    }
    
    // Sort by priority (High > Normal > Low) then by date (newest first)
    const priorityOrder = { 'High': 3, 'Normal': 2, 'Low': 1 };
    announcements.sort((a, b) => {
      const pDiff = (priorityOrder[b.priority] || 2) - (priorityOrder[a.priority] || 2);
      if (pDiff !== 0) return pDiff;
      return b.annId.localeCompare(a.annId);
    });
    
    return { success: true, announcements };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deleteAnnouncement(annId, deleterRole) {
  try {
    if (deleterRole !== 'Admin' && deleterRole !== 'Super Admin') {
      return { success: false, message: 'Access denied.' };
    }
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Announcements");
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(annId)) {
          sheet.deleteRow(i + 1);
          _invalidateCache("Announcements");
          return { success: true, message: 'Announcement deleted.' };
        }
      }
      return { success: false, message: 'Announcement not found.' };
    } finally {
      _releaseLock(lock);
    }
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── REGULARIZATION REQUESTS ─────────────────────────────────────────────────
// Sheet: "Regularization" — columns: RegID | EmpCode | EmpName | Date | MissedType | Reason | Status | AppliedDate | ApprovedBy | ApprovalDate

function submitRegularization(empCode, empName, date, missedType, reason) {
  try {
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Regularization");
      const regId = 'REG' + new Date().getTime();
      sheet.appendRow([
        regId,
        empCode,
        empName,
        date,
        missedType, // 'Check-In' or 'Check-Out'
        reason,
        'Pending',
        _today(),
        '',
        ''
      ]);
      _writeAuditLog(empCode, 'SUBMIT_REGULARIZATION', 'Date: ' + date + ', Type: ' + missedType, '');
      
      return { success: true, message: 'Regularization request submitted.', regId };
    } finally {
      _releaseLock(lock);
    }
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getRegularizations(empCode, role, filterStatus) {
  try {
    const data = _sheetData("Regularization");
    const records = [];
    
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const rowEmp = String(data[i][1]);
      const rowStatus = String(data[i][6] || '');
      
      if (role === 'Employee' && rowEmp !== String(empCode)) continue;
      if (filterStatus && filterStatus !== 'All' && rowStatus !== filterStatus) continue;
      
      records.push({
        regId: String(data[i][0]),
        empCode: String(data[i][1]),
        empName: String(data[i][2]),
        date: _formatDate(data[i][3]),
        missedType: String(data[i][4]),
        reason: String(data[i][5]),
        status: rowStatus,
        appliedDate: _formatDate(data[i][7]),
        approvedBy: String(data[i][8] || ''),
        approvalDate: _formatDate(data[i][9])
      });
    }
    
    records.sort((a, b) => b.regId.localeCompare(a.regId));
    return { success: true, records };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function approveRejectRegularization(regId, action, approverName, approverCode) {
  try {
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Regularization");
      const data = sheet.getDataRange().getValues();
      let targetRow = -1, regRow = null;
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(regId)) {
          targetRow = i + 1;
          regRow = data[i];
          break;
        }
      }
      if (targetRow === -1) return { success: false, message: 'Regularization not found.' };
      if (String(regRow[6]) !== 'Pending') return { success: false, message: 'Already ' + regRow[6] + '.' };
      
      const newStatus = action === 'Approve' ? 'Approved' : 'Rejected';
      sheet.getRange(targetRow, 7).setValue(newStatus);
      sheet.getRange(targetRow, 9).setValue(approverName);
      sheet.getRange(targetRow, 10).setValue(_today());
      
      // If approved, auto check-in/out for that date
      if (newStatus === 'Approved') {
        try {
          const attSheet = _getSheet("Attendance Sheet");
          const attData = attSheet.getDataRange().getValues();
          const regDate = _formatDate(regRow[3]);
          const regEmp  = String(regRow[1]);
          const regEmpName = String(regRow[2]);
          const missedType = String(regRow[4]);

          // Default regularization times
          const DEFAULT_CHECK_IN  = '09:00:00';
          const DEFAULT_CHECK_OUT = '18:00:00';

          /** Recalculate working hours string from two HH:mm:ss strings */
          function _calcWorkingHours(checkInStr, checkOutStr) {
            try {
              const [inH, inM, inS]   = checkInStr.split(':').map(Number);
              const [outH, outM, outS] = checkOutStr.split(':').map(Number);
              let totalSecs = (outH * 3600 + outM * 60 + outS) - (inH * 3600 + inM * 60 + inS);
              if (totalSecs < 0) totalSecs += 24 * 3600; // overnight
              const hrs  = Math.floor(totalSecs / 3600);
              const mins = Math.floor((totalSecs % 3600) / 60);
              return hrs + 'h ' + String(mins).padStart(2, '0') + 'm';
            } catch(e) { return ''; }
          }

          // Try to find an existing attendance row for this employee + date
          let foundRow = -1;
          for (let i = 1; i < attData.length; i++) {
            if (String(attData[i][1]) === regEmp && _formatDate(attData[i][2]) === regDate) {
              foundRow = i + 1; // 1-based sheet row
              break;
            }
          }

          if (foundRow !== -1) {
            // ── Row exists — patch the missing field ──────────────────
            const rowIdx = foundRow - 1; // 0-based array index
            let checkIn  = attData[rowIdx][3] ? _formatTime(attData[rowIdx][3]) : '';
            let checkOut = attData[rowIdx][4] ? _formatTime(attData[rowIdx][4]) : '';

            if (missedType === 'Check-In') {
              checkIn = DEFAULT_CHECK_IN;
              attSheet.getRange(foundRow, 4).setValue(checkIn);   // Col D — Check-In
              attSheet.getRange(foundRow, 8).setValue('Present'); // Col H — Status
              attSheet.getRange(foundRow, 9).setValue('On Time'); // Col I — Marked_As
            } else if (missedType === 'Check-Out') {
              checkOut = DEFAULT_CHECK_OUT;
              attSheet.getRange(foundRow, 5).setValue(checkOut);  // Col E — Check-Out
            }

            // Recalculate working hours whenever both times are now available
            if (checkIn && checkOut) {
              const wh = _calcWorkingHours(checkIn, checkOut);
              attSheet.getRange(foundRow, 10).setValue(wh);       // Col J — Working_Hours
            }

            attSheet.getRange(foundRow, 7).setValue('Regularized'); // Col G — Remark

          } else {
            // ── No row exists — create a new attendance record ────────
            // Get employee's shift to determine status
            let shiftId = '';
            try {
              const empData = _sheetData("Employee Master");
              const h = empData[0];
              const idCol    = _colIdx(h, "Emp ID");
              const shiftCol = _colIdx(h, "Shift ID");
              for (let i = 1; i < empData.length; i++) {
                if (String(empData[i][idCol]) === regEmp) {
                  shiftId = shiftCol >= 0 ? String(empData[i][shiftCol] || '').trim() : '';
                  break;
                }
              }
            } catch(e) {}

            let checkIn  = '';
            let checkOut = '';

            if (missedType === 'Check-In') {
              checkIn = DEFAULT_CHECK_IN;
            } else if (missedType === 'Check-Out') {
              // No check-in row at all — set both so the record is complete
              checkIn  = shiftId === 'SH_02' ? '20:00:00' : DEFAULT_CHECK_IN;
              checkOut = DEFAULT_CHECK_OUT;
            }

            const wh    = (checkIn && checkOut) ? _calcWorkingHours(checkIn, checkOut) : '';
            const attId = 'ATT' + new Date().getTime();

            // Columns: A=AttID | B=EmpID | C=Date | D=CheckIn | E=CheckOut | F=Location | G=Remark | H=Status | I=Marked_As | J=Working_Hours
            attSheet.appendRow([
              attId,
              regEmp,
              regDate,
              checkIn,
              checkOut,
              '',                // Location — not available for regularization
              'Regularized',     // Remark
              'Present',         // Status
              'On Time',         // Marked_As
              wh                 // Working_Hours
            ]);
          }
        } catch(e) {
          // Attendance update failure should not block the approval response
          Logger.log('Regularization attendance update error: ' + e.message);
        }
      }
      
      _writeAuditLog(approverCode, 'REGULARIZATION_' + newStatus.toUpperCase(), 'RegID: ' + regId, '');
      return { success: true, message: 'Regularization ' + newStatus.toLowerCase() + '.' };
    } finally {
      _releaseLock(lock);
    }
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── PAYSLIP GENERATION ──────────────────────────────────────────────────────

function generatePayslip(empCode, month, year) {
  try {
    // Get employee details
    const empData = _sheetData("Employee Master");
    const h = empData[0];
    const idCol    = _colIdx(h, "Emp ID");
    const nameCol  = _colIdx(h, "Employee Name");
    const levelCol = _colIdx(h, "Level");
    const procCol  = _colIdx(h, "Process");
    
    let emp = null;
    for (let i = 1; i < empData.length; i++) {
      if (String(empData[i][idCol]) === String(empCode)) {
        emp = {
          code:  String(empData[i][idCol]),
          name:  String(empData[i][nameCol]  || ''),
          dept:  String(empData[i][levelCol] || ''),
          desig: String(empData[i][procCol]  || '')
        };
        break;
      }
    }
    if (!emp) return { success: false, message: 'Employee not found.' };
    
    // Get salary details
    const salData = _sheetData("Salary Structure");
    let salary = null;
    for (let i = 1; i < salData.length; i++) {
      if (String(salData[i][0]) === String(empCode)) {
        // A(0):Emp_ID | B(1):Type Of Wages | C(2):Basic | D(3):H.R.A | E(4):Special E.P.F |
        // F(5):CCA | G(6):Conveyance Allowance | H(7):Medical Allowance | I(8):LTA |
        // J(9):Other Allowance | K(10):Gross
        salary = {
          typeOfWages: String(salData[i][1]  || ''),
          basic:       Number(salData[i][2])  || 0,
          hra:         Number(salData[i][3])  || 0,
          specialEpf:  Number(salData[i][4])  || 0,
          cca:         Number(salData[i][5])  || 0,
          conveyance:  Number(salData[i][6])  || 0,
          medical:     Number(salData[i][7])  || 0,
          lta:         Number(salData[i][8])  || 0,
          other:       Number(salData[i][9])  || 0,
          gross:       Number(salData[i][10]) || 0
        };
        break;
      }
    }
    if (!salary) return { success: false, message: 'Salary record not found.' };
    
    // Generate HTML payslip
    const fmt = n => '₹' + Number(n).toLocaleString('en-IN');
    const html = '<html><head><style>' +
      'body { font-family: Arial, sans-serif; margin: 40px; color: #2d3436; }' +
      'h1 { text-align: center; color: #ff4757; font-size: 1.4rem; margin-bottom: 6px; }' +
      '.sub { text-align:center; color:#636e72; font-size:.8rem; margin-bottom:20px; }' +
      '.emp-info { background:#f0f2f5; border-radius:8px; padding:14px 18px; margin-bottom:20px; font-size:.85rem; }' +
      '.emp-info p { margin:4px 0; }' +
      'table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size:.85rem; }' +
      'th, td { border: 1px solid #ddd; padding: 9px 12px; text-align: left; }' +
      'th { background: #f0f2f5; font-weight: bold; color:#2d3436; }' +
      'tr.gross-row td { background:#ff4757; color:#fff; font-weight:bold; }' +
      '.footer { margin-top:28px; font-size:.72rem; color:#999; text-align:center; }' +
      '</style></head><body>' +
      '<h1>Salary Slip</h1>' +
      '<p class="sub">' + month + ' / ' + year + '</p>' +
      '<div class="emp-info">' +
      '<p><strong>Employee:</strong> ' + emp.name + ' &nbsp;|&nbsp; <strong>ID:</strong> ' + emp.code + '</p>' +
      '<p><strong>Level:</strong> ' + emp.dept + ' &nbsp;|&nbsp; <strong>Process:</strong> ' + emp.desig + '</p>' +
      (salary.typeOfWages ? '<p><strong>Type of Wages:</strong> ' + salary.typeOfWages + '</p>' : '') +
      '</div>' +
      '<table>' +
      '<tr><th>Component</th><th>Amount</th></tr>' +
      '<tr><td>Basic</td><td>' + fmt(salary.basic) + '</td></tr>' +
      '<tr><td>H.R.A</td><td>' + fmt(salary.hra) + '</td></tr>' +
      '<tr><td>Special E.P.F</td><td>' + fmt(salary.specialEpf) + '</td></tr>' +
      '<tr><td>CCA</td><td>' + fmt(salary.cca) + '</td></tr>' +
      '<tr><td>Conveyance Allowance</td><td>' + fmt(salary.conveyance) + '</td></tr>' +
      '<tr><td>Medical Allowance</td><td>' + fmt(salary.medical) + '</td></tr>' +
      '<tr><td>LTA</td><td>' + fmt(salary.lta) + '</td></tr>' +
      '<tr><td>Other Allowance</td><td>' + fmt(salary.other) + '</td></tr>' +
      '<tr class="gross-row"><td>Gross</td><td>' + fmt(salary.gross) + '</td></tr>' +
      '</table>' +
      '<p class="footer">Generated on ' + _today() + ' &nbsp;|&nbsp; Workforce Operations Portal</p>' +
      '</body></html>';
    
    return { success: true, html };
  } catch(e) {
    return { success: false, message: e.message };
  }
}


// ─── REPORTS & EXPORTS ───────────────────────────────────────────────────────

function getAttendanceReport(month, year, requestorRole) {
  try {
    if (requestorRole !== 'Admin' && requestorRole !== 'Super Admin' && requestorRole !== 'Manager') {
      return { success: false, message: 'Access denied.' };
    }

    // Always read fresh — never use cache for reports
    const attSheet = _getSheet("Attendance Sheet");
    const attData  = attSheet.getDataRange().getValues();
    const empData  = _sheetData("Employee Master");

    // ── Attendance sheet column indices (positional — matches actual sheet layout) ──
    // A=AttID | B=EmpID | C=Date | D=CheckIn | E=CheckOut | F=Location | G=Availability | H=Status | I=Marked_As | J=Working_Hours
    const empCol    = 1;
    const dateCol   = 2;
    const inCol     = 3;
    const statusCol = 7;
    const markCol   = 8;

    // ── Employee Master column indices ──
    const eH      = empData[0];
    const idCol   = _colIdx(eH, 'Emp ID');
    const nameCol = _colIdx(eH, 'Employee Name');
    const deptCol = _colIdx(eH, 'Level');

    // Build emp lookup map
    const empMap = {};
    for (let i = 1; i < empData.length; i++) {
      const code = String(empData[i][idCol] || '').trim();
      if (code) {
        empMap[code] = {
          name: String(empData[i][nameCol] || ''),
          dept: String(empData[i][deptCol] || '')
        };
      }
    }

    const targetMonthYear = String(month).padStart(2, '0') + '/' + year;
    const summary = {}; // empCode → { present, absent, late, halfDay }

    for (let i = 1; i < attData.length; i++) {
      const rowEmp = String(attData[i][empCol] || '').trim();
      if (!rowEmp) continue;

      // ── Date matching ──
      let rowMonthYear = '';
      const rawDate = attData[i][dateCol];
      if (rawDate instanceof Date) {
        if (rawDate.getFullYear() >= 1900) {
          rowMonthYear = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'MM/yyyy');
        }
      } else {
        const parts = String(rawDate).trim().split('/');
        if (parts.length === 3) {
          // dd/MM/yyyy → MM/yyyy
          rowMonthYear = parts[1].padStart(2, '0') + '/' + parts[2];
        }
      }
      if (rowMonthYear !== targetMonthYear) continue;

      if (!summary[rowEmp]) summary[rowEmp] = { present: 0, absent: 0, late: 0, halfDay: 0 };

      // Read Status (H) and Marked_As (I) — normalise whitespace and case
      const st       = String(attData[i][statusCol] || '').toLowerCase().replace(/\s+/g, ' ').trim();
      const markedAs = String(attData[i][markCol]   || '').toLowerCase().replace(/\s+/g, ' ').trim();

      if (st === 'absent') {
        summary[rowEmp].absent++;
      } else if (st === 'leave') {
        // leave days — skip (not present/absent)
      } else if (markedAs.indexOf('half') !== -1) {
        // "Half Day" — counts as present but flagged
        summary[rowEmp].halfDay++;
        summary[rowEmp].present++;
      } else if (markedAs.indexOf('late') !== -1) {
        // "Late Arrival"
        summary[rowEmp].present++;
        summary[rowEmp].late++;
      } else if (attData[i][inCol]) {
        // Has a check-in time → On Time present
        summary[rowEmp].present++;
      }
    }

    const rows = [];
    for (const empCode in summary) {
      const emp = empMap[empCode] || { name: empCode, dept: '' };
      rows.push({ empCode, name: emp.name, dept: emp.dept, ...summary[empCode] });
    }
    rows.sort((a, b) => a.name.localeCompare(b.name));

    return { success: true, rows, month: targetMonthYear };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getLeaveReport(month, year, requestorRole) {
  try {
    if (requestorRole !== 'Admin' && requestorRole !== 'Super Admin' && requestorRole !== 'Manager') {
      return { success: false, message: 'Access denied.' };
    }
    const lrData = _sheetData("Leave Request");
    const rows = [];
    
    for (let i = 1; i < lrData.length; i++) {
      if (!lrData[i][0]) continue;
      const fromDate = lrData[i][4];
      let rowMonth = '', rowYear = '';
      if (fromDate instanceof Date) {
        rowMonth = String(fromDate.getMonth() + 1).padStart(2, '0');
        rowYear = String(fromDate.getFullYear());
      } else {
        const s = String(fromDate || '');
        if (s.indexOf('-') !== -1) {
          const parts = s.split('-');
          rowYear = parts[0]; rowMonth = parts[1];
        } else if (s.indexOf('/') !== -1) {
          const parts = s.split('/');
          rowMonth = parts[1]; rowYear = parts[2];
        }
      }
      
      if (String(month).padStart(2, '0') !== rowMonth || String(year) !== rowYear) continue;
      
      rows.push({
        leaveId: String(lrData[i][0]),
        empCode: String(lrData[i][1]),
        empName: String(lrData[i][2]),
        leaveType: String(lrData[i][3]),
        fromDate: _formatDate(lrData[i][4]),
        toDate: _formatDate(lrData[i][5]),
        days: lrData[i][6],
        status: String(lrData[i][8] || '')
      });
    }
    
    return { success: true, rows };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── UPDATE PERSONAL DETAILS (Employee self-service) ─────────────────────────

function updatePersonalDetails(empCode, fields) {
  try {
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Employee Master");
      const data  = sheet.getDataRange().getValues();
      const h     = data[0];
      const persMobCol   = _colIdx(h, "Personal Mobile Number");
      const persEmailCol = _colIdx(h, "Personal Email ID");
      const addressCol   = _colIdx(h, "Permanent Address");
      const cityCol      = _colIdx(h, "City");
      const districtCol  = _colIdx(h, "District");
      const stateCol     = _colIdx(h, "State");

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(empCode)) {
          const row = i + 1;
          if (fields.personalMobile !== undefined && persMobCol   >= 0) sheet.getRange(row, persMobCol   + 1).setValue(fields.personalMobile);
          if (fields.personalEmail  !== undefined && persEmailCol >= 0) sheet.getRange(row, persEmailCol + 1).setValue(fields.personalEmail);
          if (fields.address        !== undefined && addressCol   >= 0) sheet.getRange(row, addressCol   + 1).setValue(fields.address);
          if (fields.city           !== undefined && cityCol      >= 0) sheet.getRange(row, cityCol      + 1).setValue(fields.city);
          if (fields.district       !== undefined && districtCol  >= 0) sheet.getRange(row, districtCol  + 1).setValue(fields.district);
          if (fields.state          !== undefined && stateCol     >= 0) sheet.getRange(row, stateCol     + 1).setValue(fields.state);
          _invalidateCache("Employee Master");
          _writeAuditLog(empCode, 'PERSONAL_DETAILS_UPDATED', 'Self-service update', '');
          return { success: true, message: 'Personal details updated successfully.' };
        }
      }
      return { success: false, message: 'Employee not found.' };
    } finally {
      _releaseLock(lock);
    }
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── TASK STATUS UPDATE ───────────────────────────────────────────────────────

function updateTaskStatus(docId, newStatus, empCode) {
  try {
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Task Sheet");
      const data  = sheet.getDataRange().getValues();
      const h     = data[0];
      const docIdCol = h.findIndex(c => String(c).toLowerCase().trim() === 'doc id');
      const statusCol = h.findIndex(c => String(c).toLowerCase().trim() === 'status');
      
      if (statusCol === -1) return { success: false, message: 'Status column not found in Task Sheet.' };
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][docIdCol]) === String(docId)) {
          sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
          _writeAuditLog(empCode, 'TASK_STATUS_UPDATED', 'DocID: ' + docId + ' → ' + newStatus, '');
          return { success: true, message: 'Task status updated to ' + newStatus };
        }
      }
      return { success: false, message: 'Task not found.' };
    } finally {
      _releaseLock(lock);
    }
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── NOTIFICATION COUNTS ─────────────────────────────────────────────────────
/**
 * Returns unread/new counts for every notification category.
 * Called by the frontend every 60 s to drive the badge system.
 *
 * Return shape:
 * {
 *   announcements:           { count, sub, ts, items[] },
 *   holidays:                { count, sub, ts, items[] },
 *   leave:                   { count, sub, ts, items[] },
 *   gatepass:                { count, sub, ts, items[] },
 *   tasks:                   { count, sub, ts, items[] },
 *   attendance:              { count, sub, ts, items[] },
 *   regularization:          { count, sub, ts, items[] },
 *   policies:                { count, sub, ts, items[] },
 *   leaveApprovals:          { count, sub, ts, items[] },   // manager/admin only
 *   gatepassApprovals:       { count, sub, ts, items[] },   // manager/admin only
 *   regularizationApprovals: { count, sub, ts, items[] }    // manager/admin only
 * }
 */
function getNotificationCounts(empCode, role) {
  try {
    const result = {};
    const today  = new Date();
    const isManagerOrAdmin = (role === 'Manager' || role === 'Admin' || role === 'Super Admin');

    // ── 1. ANNOUNCEMENTS — count active (non-expired) announcements ──────────
    try {
      const annSheet = _getSpreadsheet().getSheetByName('Announcements');
      if (annSheet) {
        const annData = annSheet.getDataRange().getValues();
        const annItems = [];
        for (let i = 1; i < annData.length; i++) {
          if (!annData[i][0]) continue;
          // Expiry check
          const expiry = annData[i][6];
          if (expiry) {
            let expiryDate;
            if (expiry instanceof Date) {
              if (expiry.getFullYear() >= 1970) expiryDate = expiry;
            } else {
              const s = String(expiry).trim();
              if (s) {
                const parts = s.split('/');
                if (parts.length === 3) expiryDate = new Date(Number(parts[2]), Number(parts[1])-1, Number(parts[0]));
              }
            }
            if (expiryDate && !isNaN(expiryDate.getTime())) {
              expiryDate.setHours(23,59,59,999);
              if (expiryDate < today) continue;
            }
          }
          annItems.push({
            title: String(annData[i][1] || 'Announcement'),
            sub:   String(annData[i][3] || '') + (annData[i][5] ? ' · ' + _formatDate(annData[i][5]) : ''),
            view:  'announcements',
            ts:    annData[i][5] instanceof Date ? annData[i][5].getTime() : 0
          });
        }
        result.announcements = {
          count: annItems.length,
          sub:   annItems.length + ' active announcement' + (annItems.length !== 1 ? 's' : ''),
          ts:    annItems.length ? annItems[0].ts : 0,
          items: annItems.slice(0, 10)
        };
      }
    } catch(e) { result.announcements = { count: 0, items: [] }; }

    // ── 2. HOLIDAYS — upcoming holidays in next 30 days ──────────────────────
    try {
      const holData = _sheetData('Holiday List');
      const holItems = [];
      const in30 = new Date(today.getTime() + 30 * 24 * 60 * 60 * 1000);
      for (let i = 1; i < holData.length; i++) {
        if (!holData[i][0]) continue;
        let hDate;
        const raw = holData[i][1];
        if (raw instanceof Date) {
          hDate = raw;
        } else {
          const s = String(raw || '').trim();
          const parts = s.split('/');
          if (parts.length === 3) hDate = new Date(Number(parts[2]), Number(parts[1])-1, Number(parts[0]));
          else hDate = new Date(s);
        }
        if (!hDate || isNaN(hDate.getTime())) continue;
        if (hDate >= today && hDate <= in30) {
          holItems.push({
            title: String(holData[i][2] || 'Holiday'),
            sub:   _formatDate(holData[i][1]) + ' · ' + String(holData[i][3] || 'Public'),
            view:  'holidays',
            ts:    hDate.getTime()
          });
        }
      }
      holItems.sort((a,b) => a.ts - b.ts);
      result.holidays = {
        count: holItems.length,
        sub:   holItems.length + ' upcoming holiday' + (holItems.length !== 1 ? 's' : '') + ' in 30 days',
        ts:    holItems.length ? holItems[0].ts : 0,
        items: holItems
      };
    } catch(e) { result.holidays = { count: 0, items: [] }; }

    // ── 3. LEAVE — employee's pending leave requests ──────────────────────────
    try {
      const lrData = _sheetData('Leave Request');
      const leaveItems = [];
      for (let i = 1; i < lrData.length; i++) {
        if (!lrData[i][0]) continue;
        if (String(lrData[i][1]).trim() !== String(empCode).trim()) continue;
        const status = String(lrData[i][8] || '');
        if (status === 'Pending') {
          leaveItems.push({
            title: 'Leave Request Pending — ' + String(lrData[i][3] || ''),
            sub:   _formatDate(lrData[i][4]) + ' to ' + _formatDate(lrData[i][5]),
            view:  'leave',
            ts:    0
          });
        } else if (status === 'Approved' || status === 'Rejected') {
          // Show recently actioned (within last 7 days)
          const approvalDate = lrData[i][11];
          let aDate;
          if (approvalDate instanceof Date) aDate = approvalDate;
          else {
            const s = String(approvalDate || '').trim();
            const parts = s.split('/');
            if (parts.length === 3) aDate = new Date(Number(parts[2]), Number(parts[1])-1, Number(parts[0]));
          }
          if (aDate && !isNaN(aDate.getTime())) {
            const diffDays = (today.getTime() - aDate.getTime()) / (1000*60*60*24);
            if (diffDays <= 7) {
              leaveItems.push({
                title: 'Leave ' + status + ' — ' + String(lrData[i][3] || ''),
                sub:   'By ' + String(lrData[i][10] || '') + ' · ' + _formatDate(lrData[i][11]),
                view:  'leave',
                ts:    aDate.getTime()
              });
            }
          }
        }
      }
      result.leave = {
        count: leaveItems.length,
        sub:   leaveItems.length + ' leave notification' + (leaveItems.length !== 1 ? 's' : ''),
        ts:    leaveItems.length ? Math.max(...leaveItems.map(x=>x.ts)) : 0,
        items: leaveItems.slice(0, 10)
      };
    } catch(e) { result.leave = { count: 0, items: [] }; }

    // ── 4. GATE PASS — employee's pending / recently actioned ────────────────
    try {
      const gpData = _sheetData('Gate Pass');
      const gpItems = [];
      for (let i = 1; i < gpData.length; i++) {
        if (!gpData[i][0]) continue;
        if (String(gpData[i][1]).trim() !== String(empCode).trim()) continue;
        const status = String(gpData[i][8] || '');
        if (status === 'Pending') {
          gpItems.push({
            title: 'Gate Pass Pending — ' + String(gpData[i][3] || ''),
            sub:   'Applied ' + _formatDate(gpData[i][9]),
            view:  'gatepass',
            ts:    0
          });
        } else if (status === 'Approved' || status === 'Rejected') {
          const approvalDate = gpData[i][11];
          let aDate;
          if (approvalDate instanceof Date) aDate = approvalDate;
          else {
            const s = String(approvalDate || '').trim();
            const parts = s.split('/');
            if (parts.length === 3) aDate = new Date(Number(parts[2]), Number(parts[1])-1, Number(parts[0]));
          }
          if (aDate && !isNaN(aDate.getTime())) {
            const diffDays = (today.getTime() - aDate.getTime()) / (1000*60*60*24);
            if (diffDays <= 7) {
              gpItems.push({
                title: 'Gate Pass ' + status + ' — ' + String(gpData[i][3] || ''),
                sub:   'By ' + String(gpData[i][10] || '') + ' · ' + _formatDate(gpData[i][11]),
                view:  'gatepass',
                ts:    aDate.getTime()
              });
            }
          }
        }
      }
      result.gatepass = {
        count: gpItems.length,
        sub:   gpItems.length + ' gate pass notification' + (gpItems.length !== 1 ? 's' : ''),
        ts:    gpItems.length ? Math.max(...gpItems.map(x=>x.ts)) : 0,
        items: gpItems.slice(0, 10)
      };
    } catch(e) { result.gatepass = { count: 0, items: [] }; }

    // ── 5. TASKS — employee's pending / overdue tasks ─────────────────────────
    try {
      const taskData = _sheetData('Task Sheet');
      const taskItems = [];
      if (taskData.length > 1) {
        const h = taskData[0];
        const col = name => {
          const n = name.toLowerCase().trim();
          const idx = h.findIndex(c => String(c).toLowerCase().trim() === n);
          return idx >= 0 ? idx : -1;
        };
        const docIdCol    = col('doc id')      >= 0 ? col('doc id')      : 0;
        const blogCol     = col('blog name')   >= 0 ? col('blog name')   : 1;
        const deadlineCol = col('deadline')    >= 0 ? col('deadline')    : 2;
        const doerCol     = col('doer name')   >= 0 ? col('doer name')   : 4;
        const statusCol   = col('status')      >= 0 ? col('status')      : -1;

        // Get employee full name from Employee Master for name-based matching
        let empFullName = '';
        try {
          const empData = _sheetData('Employee Master');
          const eh = empData[0];
          const idIdx   = _colIdx(eh, 'Emp ID');
          const nameIdx = _colIdx(eh, 'Employee Name');
          for (let i = 1; i < empData.length; i++) {
            if (String(empData[i][idIdx]).trim() === String(empCode).trim()) {
              empFullName = String(empData[i][nameIdx] || '').trim();
              break;
            }
          }
        } catch(e) {}

        const empCodeLower    = String(empCode).trim().toLowerCase();
        const empNameLower    = empFullName.toLowerCase();

        for (let i = 1; i < taskData.length; i++) {
          if (!taskData[i][docIdCol] && !taskData[i][blogCol]) continue;

          const rowDoer   = String(taskData[i][doerCol] || '').trim();
          const dashIdx   = rowDoer.lastIndexOf(' - ');
          const doerEmpId = dashIdx >= 0 ? rowDoer.substring(dashIdx + 3).trim().toLowerCase() : '';
          const doerName  = dashIdx >= 0 ? rowDoer.substring(0, dashIdx).trim().toLowerCase() : rowDoer.toLowerCase();

          // Match ONLY by emp code or emp name — no email matching
          const matched =
            (empCodeLower && doerEmpId && doerEmpId === empCodeLower) ||
            (empNameLower && doerName  && doerName  === empNameLower);
          if (!matched) continue;

          const taskStatus = statusCol >= 0 ? String(taskData[i][statusCol] || '') : '';
          if (taskStatus === 'Completed') continue;

          // Check if overdue
          const rawDl = taskData[i][deadlineCol];
          let dlDate;
          if (rawDl instanceof Date) dlDate = rawDl;
          else {
            const s = String(rawDl || '').trim();
            if (/^\d{2}\/\d{2}\/\d{4}/.test(s)) {
              const [d,m,y] = s.split('/');
              dlDate = new Date(Number(y), Number(m)-1, Number(d));
            } else dlDate = new Date(s);
          }
          const isOverdue = dlDate && !isNaN(dlDate.getTime()) && dlDate < today;

          // Sub-line: emp code + name + deadline
          const empLabel      = empFullName ? empCode + ' · ' + empFullName : empCode;
          const deadlineLabel = rawDl ? _formatDate(rawDl) : '—';

          taskItems.push({
            title: (isOverdue ? '🚨 OVERDUE: ' : '') + String(taskData[i][blogCol] || 'Task'),
            sub:   empLabel + ' · Deadline: ' + deadlineLabel,
            view:  'tasks',
            ts:    dlDate && !isNaN(dlDate.getTime()) ? dlDate.getTime() : 0
          });
        }
      }
      result.tasks = {
        count: taskItems.length,
        sub:   taskItems.length + ' pending task' + (taskItems.length !== 1 ? 's' : ''),
        ts:    taskItems.length ? Math.max(...taskItems.map(x=>x.ts)) : 0,
        items: taskItems.slice(0, 10)
      };
    } catch(e) { result.tasks = { count: 0, items: [] }; }

    // ── 6. ATTENDANCE — check if today's check-in is missing ─────────────────
    try {
      const attData = _sheetData('Attendance Sheet');
      const todayStr = _today();
      let checkedInToday = false;
      let checkedOutToday = false;
      for (let i = attData.length - 1; i >= 1; i--) {
        if (String(attData[i][1]).trim() !== String(empCode).trim()) continue;
        if (_formatDate(attData[i][2]) === todayStr) {
          checkedInToday  = !!attData[i][3];
          checkedOutToday = !!attData[i][4];
          break;
        }
      }
      const attItems = [];
      const nowHour = new Date().getHours();
      // Alert: after 9 AM and not checked in yet
      if (!checkedInToday && nowHour >= 9 && nowHour < 20) {
        attItems.push({ title: 'You have not checked in today', sub: 'Tap to mark attendance', view: 'attendance', ts: Date.now() });
      }
      // Alert: after 6 PM and checked in but not checked out
      if (checkedInToday && !checkedOutToday && nowHour >= 18) {
        attItems.push({ title: 'You have not checked out today', sub: 'Tap to mark checkout', view: 'attendance', ts: Date.now() });
      }
      result.attendance = {
        count: attItems.length,
        sub:   attItems.length ? attItems[0].title : '',
        ts:    Date.now(),
        items: attItems
      };
    } catch(e) { result.attendance = { count: 0, items: [] }; }

    // ── 7. REGULARIZATION — employee's pending regularizations ───────────────
    try {
      const regData = _sheetData('Regularization');
      const regItems = [];
      for (let i = 1; i < regData.length; i++) {
        if (!regData[i][0]) continue;
        if (String(regData[i][1]).trim() !== String(empCode).trim()) continue;
        const status = String(regData[i][6] || '');
        if (status === 'Pending') {
          regItems.push({
            title: 'Regularization Pending — ' + String(regData[i][4] || ''),
            sub:   'Date: ' + _formatDate(regData[i][3]),
            view:  'regularization',
            ts:    0
          });
        } else if ((status === 'Approved' || status === 'Rejected')) {
          const aDate = regData[i][9] instanceof Date ? regData[i][9] : null;
          if (aDate) {
            const diffDays = (today.getTime() - aDate.getTime()) / (1000*60*60*24);
            if (diffDays <= 7) {
              regItems.push({
                title: 'Regularization ' + status,
                sub:   'By ' + String(regData[i][8] || '') + ' · ' + _formatDate(regData[i][9]),
                view:  'regularization',
                ts:    aDate.getTime()
              });
            }
          }
        }
      }
      result.regularization = {
        count: regItems.length,
        sub:   regItems.length + ' regularization notification' + (regItems.length !== 1 ? 's' : ''),
        ts:    regItems.length ? Math.max(...regItems.map(x=>x.ts)) : 0,
        items: regItems.slice(0, 10)
      };
    } catch(e) { result.regularization = { count: 0, items: [] }; }

    // ── 8. POLICIES — new policies in last 7 days ─────────────────────────────
    try {
      const polData = _sheetData('Policies');
      const polItems = [];
      for (let i = 1; i < polData.length; i++) {
        if (!polData[i][0]) continue;
        const uploadDate = polData[i][10];
        let uDate;
        if (uploadDate instanceof Date) uDate = uploadDate;
        else {
          const s = String(uploadDate || '').trim();
          const parts = s.split('/');
          if (parts.length === 3) uDate = new Date(Number(parts[2]), Number(parts[1])-1, Number(parts[0]));
        }
        if (!uDate || isNaN(uDate.getTime())) continue;
        const diffDays = (today.getTime() - uDate.getTime()) / (1000*60*60*24);
        if (diffDays <= 7) {
          polItems.push({
            title: String(polData[i][1] || 'New Policy'),
            sub:   String(polData[i][2] || '') + ' · ' + _formatDate(polData[i][10]),
            view:  'policies',
            ts:    uDate.getTime()
          });
        }
      }
      result.policies = {
        count: polItems.length,
        sub:   polItems.length + ' new polic' + (polItems.length !== 1 ? 'ies' : 'y') + ' this week',
        ts:    polItems.length ? Math.max(...polItems.map(x=>x.ts)) : 0,
        items: polItems.slice(0, 10)
      };
    } catch(e) { result.policies = { count: 0, items: [] }; }

    // ── 9–11. APPROVAL QUEUES (manager / admin only) ──────────────────────────
    if (isManagerOrAdmin) {
      // Leave approvals
      try {
        const lrData = _sheetData('Leave Request');
        const laItems = [];

        // For Manager role: build empCode→managerCode map once (avoids N+1 sheet reads)
        let empMgrMap = null;
        if (role === 'Manager') {
          empMgrMap = {};
          const empData = _sheetData('Employee Master');
          const h = empData[0];
          const idCol  = _colIdx(h, 'Emp ID');
          const mgrCol = _colIdx(h, 'Manager ID');
          for (let j = 1; j < empData.length; j++) {
            if (empData[j][idCol]) {
              empMgrMap[String(empData[j][idCol])] = String(empData[j][mgrCol] || '');
            }
          }
        }

        for (let i = 1; i < lrData.length; i++) {
          if (!lrData[i][0]) continue;
          if (String(lrData[i][8] || '').toLowerCase() !== 'pending') continue;
          if (role === 'Manager') {
            if (empMgrMap[String(lrData[i][1])] !== String(empCode)) continue;
          }
          laItems.push({
            title: String(lrData[i][2] || '') + ' — ' + String(lrData[i][3] || '') + ' Leave',
            sub:   _formatDate(lrData[i][4]) + ' to ' + _formatDate(lrData[i][5]) + ' · ' + String(lrData[i][6] || '') + ' day(s)',
            view:  'leaveApprovals',
            ts:    0
          });
        }
        result.leaveApprovals = {
          count: laItems.length,
          sub:   laItems.length + ' pending leave approval' + (laItems.length !== 1 ? 's' : ''),
          ts:    0,
          items: laItems.slice(0, 10)
        };
      } catch(e) { result.leaveApprovals = { count: 0, items: [] }; }

      // Gate pass approvals
      try {
        const gpData = _sheetData('Gate Pass');
        const gaItems = [];
        for (let i = 1; i < gpData.length; i++) {
          if (!gpData[i][0]) continue;
          if (String(gpData[i][8] || '').toLowerCase() !== 'pending') continue;
          gaItems.push({
            title: String(gpData[i][2] || '') + ' — ' + String(gpData[i][3] || '') + ' Pass',
            sub:   'Applied ' + _formatDate(gpData[i][9]),
            view:  'gatepassApprovals',
            ts:    0
          });
        }
        result.gatepassApprovals = {
          count: gaItems.length,
          sub:   gaItems.length + ' pending gate pass approval' + (gaItems.length !== 1 ? 's' : ''),
          ts:    0,
          items: gaItems.slice(0, 10)
        };
      } catch(e) { result.gatepassApprovals = { count: 0, items: [] }; }

      // Regularization approvals
      try {
        const regData = _sheetData('Regularization');
        const raItems = [];
        for (let i = 1; i < regData.length; i++) {
          if (!regData[i][0]) continue;
          if (String(regData[i][6] || '').toLowerCase() !== 'pending') continue;
          raItems.push({
            title: String(regData[i][2] || '') + ' — ' + String(regData[i][4] || ''),
            sub:   'Date: ' + _formatDate(regData[i][3]),
            view:  'regularizationApprovals',
            ts:    0
          });
        }
        result.regularizationApprovals = {
          count: raItems.length,
          sub:   raItems.length + ' pending regularization approval' + (raItems.length !== 1 ? 's' : ''),
          ts:    0,
          items: raItems.slice(0, 10)
        };
      } catch(e) { result.regularizationApprovals = { count: 0, items: [] }; }
    } else {
      result.leaveApprovals          = { count: 0, items: [] };
      result.gatepassApprovals       = { count: 0, items: [] };
      result.regularizationApprovals = { count: 0, items: [] };
    }

    return result;
  } catch(e) {
    return { error: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// ─── NEW ADVANCED FEATURES — BLOCK 1: PAYSLIP WITH DEDUCTIONS ───────────────
// ═══════════════════════════════════════════════════════════════════════════

/**
 * generatePayslipWithDeductions
 * Calculates PF (12% of Basic), ESI (0.75% of Gross if Gross ≤ 21000),
 * Professional Tax (slab-based), and derives Net Pay.
 * Also factors in attendance: Net = Gross × (presentDays / workingDays)
 */
function generatePayslipWithDeductions(empCode, month, year) {
  try {
    // ── Employee details ──────────────────────────────────────────────────
    const empData = _sheetData("Employee Master");
    const h = empData[0];
    const idCol    = _colIdx(h, "Emp ID");
    const nameCol  = _colIdx(h, "Employee Name");
    const levelCol = _colIdx(h, "Level");
    const procCol  = _colIdx(h, "Process");
    const dojCol   = _colIdx(h, "Date of Joining");

    let emp = null;
    for (let i = 1; i < empData.length; i++) {
      if (String(empData[i][idCol]) === String(empCode)) {
        emp = {
          code:  String(empData[i][idCol]),
          name:  String(empData[i][nameCol]  || ''),
          dept:  String(empData[i][levelCol] || ''),
          desig: String(empData[i][procCol]  || ''),
          doj:   _formatDate(dojCol >= 0 ? empData[i][dojCol] : '')
        };
        break;
      }
    }
    if (!emp) return { success: false, message: 'Employee not found.' };

    // ── Salary structure ──────────────────────────────────────────────────
    const salData = _sheetData("Salary Structure");
    let sal = null;
    for (let i = 1; i < salData.length; i++) {
      if (String(salData[i][0]) === String(empCode)) {
        sal = {
          typeOfWages: String(salData[i][1]  || ''),
          basic:       Number(salData[i][2])  || 0,
          hra:         Number(salData[i][3])  || 0,
          specialEpf:  Number(salData[i][4])  || 0,
          cca:         Number(salData[i][5])  || 0,
          conveyance:  Number(salData[i][6])  || 0,
          medical:     Number(salData[i][7])  || 0,
          lta:         Number(salData[i][8])  || 0,
          other:       Number(salData[i][9])  || 0,
          gross:       Number(salData[i][10]) || 0
        };
        break;
      }
    }
    if (!sal) return { success: false, message: 'Salary record not found.' };

    // ── Attendance for the month ──────────────────────────────────────────
    const attData  = _sheetData("Attendance Sheet");
    const targetMY = String(month).padStart(2,'0') + '/' + year;
    let presentDays = 0, absentDays = 0, lateDays = 0, halfDays = 0;

    for (let i = 1; i < attData.length; i++) {
      if (String(attData[i][1]).trim() !== String(empCode).trim()) continue;
      let rowMY = '';
      const raw = attData[i][2];
      if (raw instanceof Date && raw.getFullYear() >= 1900) {
        rowMY = Utilities.formatDate(raw, Session.getScriptTimeZone(), 'MM/yyyy');
      } else {
        const parts = String(raw).trim().split('/');
        if (parts.length === 3) rowMY = parts[1].padStart(2,'0') + '/' + parts[2];
      }
      if (rowMY !== targetMY) continue;

      const st  = String(attData[i][7] || '').toLowerCase().trim();
      const mk  = String(attData[i][8] || '').toLowerCase().trim();
      if (st === 'absent') { absentDays++; }
      else if (mk.indexOf('half') !== -1) { halfDays++; presentDays += 0.5; }
      else if (mk.indexOf('late') !== -1) { lateDays++; presentDays++; }
      else if (attData[i][3]) { presentDays++; }
    }

    // Working days in the month (Mon–Sat, excluding holidays)
    const daysInMonth = new Date(Number(year), Number(month), 0).getDate();
    let workingDays = 0;
    const holData = _sheetData("Holiday List");
    const holSet  = new Set();
    for (let i = 1; i < holData.length; i++) {
      const hd = holData[i][1];
      if (hd instanceof Date) holSet.add(Utilities.formatDate(hd, Session.getScriptTimeZone(), 'dd/MM/yyyy'));
      else { const s = String(hd||'').trim(); if (s) holSet.add(s); }
    }
    for (let d = 1; d <= daysInMonth; d++) {
      const dt  = new Date(Number(year), Number(month) - 1, d);
      const dow = dt.getDay(); // 0=Sun,6=Sat
      if (dow === 0) continue; // Sunday off
      const ds = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      if (holSet.has(ds)) continue;
      workingDays++;
    }
    if (workingDays === 0) workingDays = 26; // fallback

    // ── Earnings (pro-rated) ──────────────────────────────────────────────
    const ratio       = presentDays > 0 ? Math.min(presentDays / workingDays, 1) : 0;
    const earnedGross = Math.round(sal.gross * ratio);
    const earnedBasic = Math.round(sal.basic * ratio);

    // ── Deductions ────────────────────────────────────────────────────────
    // PF: 12% of Basic (employee share)
    const pf = Math.round(earnedBasic * 0.12);
    // ESI: 0.75% of Gross if Gross ≤ 21000
    const esi = sal.gross <= 21000 ? Math.round(earnedGross * 0.0075) : 0;
    // Professional Tax (India slab — monthly gross)
    let pt = 0;
    if (earnedGross > 15000) pt = 200;
    else if (earnedGross > 10000) pt = 150;
    else if (earnedGross > 7500) pt = 75;
    // LWF (Labour Welfare Fund) — ₹25 flat
    const lwf = 25;
    const totalDeductions = pf + esi + pt + lwf;
    const netPay = Math.max(0, earnedGross - totalDeductions);

    // ── Month name ────────────────────────────────────────────────────────
    const monthNames = ['','January','February','March','April','May','June',
                        'July','August','September','October','November','December'];
    const monthName = monthNames[Number(month)] || month;

    const fmt = n => '₹' + Number(n).toLocaleString('en-IN');

    const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>
  body{font-family:Arial,sans-serif;margin:0;padding:32px;color:#2d3436;background:#fff}
  .slip-header{text-align:center;border-bottom:3px solid #ff4757;padding-bottom:16px;margin-bottom:20px}
  .slip-header h1{color:#ff4757;font-size:1.5rem;margin:0 0 4px}
  .slip-header p{color:#636e72;font-size:.8rem;margin:0}
  .emp-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px 24px;background:#f8f9fa;border-radius:8px;padding:14px 18px;margin-bottom:20px;font-size:.82rem}
  .emp-grid span{color:#636e72;font-size:.72rem;text-transform:uppercase;letter-spacing:.05em}
  .emp-grid strong{display:block;color:#2d3436;font-size:.88rem}
  .att-summary{display:flex;gap:12px;margin-bottom:20px;flex-wrap:wrap}
  .att-box{flex:1;min-width:80px;text-align:center;background:#f8f9fa;border-radius:8px;padding:10px 8px}
  .att-box .num{font-size:1.4rem;font-weight:800;color:#2d3436;font-family:monospace}
  .att-box .lbl{font-size:.65rem;color:#636e72;text-transform:uppercase;letter-spacing:.05em}
  table{width:100%;border-collapse:collapse;font-size:.82rem;margin-bottom:16px}
  th{background:#f0f2f5;padding:9px 12px;text-align:left;font-size:.7rem;text-transform:uppercase;letter-spacing:.05em;color:#636e72}
  td{padding:9px 12px;border-bottom:1px solid #f0f2f5}
  .section-title{font-size:.72rem;font-weight:800;text-transform:uppercase;letter-spacing:.08em;color:#636e72;margin:16px 0 8px;padding-bottom:4px;border-bottom:1px solid #e0e5ec}
  .net-row td{background:#ff4757;color:#fff;font-weight:800;font-size:1rem}
  .footer{margin-top:24px;font-size:.68rem;color:#b2bec3;text-align:center;border-top:1px solid #f0f2f5;padding-top:12px}
  @media print{body{padding:16px}.footer{display:none}}
</style></head><body>
<div class="slip-header">
  <h1>Rishi Seals Pvt. Ltd.</h1>
  <p>Salary Slip — ${monthName} ${year}</p>
</div>
<div class="emp-grid">
  <div><span>Employee Name</span><strong>${emp.name}</strong></div>
  <div><span>Employee ID</span><strong>${emp.code}</strong></div>
  <div><span>Level / Department</span><strong>${emp.dept || '—'}</strong></div>
  <div><span>Process</span><strong>${emp.desig || '—'}</strong></div>
  <div><span>Date of Joining</span><strong>${emp.doj || '—'}</strong></div>
  <div><span>Type of Wages</span><strong>${sal.typeOfWages || '—'}</strong></div>
</div>
<div class="section-title">Attendance Summary</div>
<div class="att-summary">
  <div class="att-box"><div class="num">${workingDays}</div><div class="lbl">Working Days</div></div>
  <div class="att-box"><div class="num">${presentDays}</div><div class="lbl">Present</div></div>
  <div class="att-box"><div class="num">${absentDays}</div><div class="lbl">Absent</div></div>
  <div class="att-box"><div class="num">${lateDays}</div><div class="lbl">Late</div></div>
  <div class="att-box"><div class="num">${halfDays}</div><div class="lbl">Half Days</div></div>
</div>
<div class="section-title">Earnings</div>
<table>
  <tr><th>Component</th><th>Monthly Rate</th><th>Earned (Pro-rated)</th></tr>
  <tr><td>Basic</td><td>${fmt(sal.basic)}</td><td>${fmt(Math.round(sal.basic*ratio))}</td></tr>
  <tr><td>H.R.A</td><td>${fmt(sal.hra)}</td><td>${fmt(Math.round(sal.hra*ratio))}</td></tr>
  <tr><td>Special E.P.F</td><td>${fmt(sal.specialEpf)}</td><td>${fmt(Math.round(sal.specialEpf*ratio))}</td></tr>
  <tr><td>CCA</td><td>${fmt(sal.cca)}</td><td>${fmt(Math.round(sal.cca*ratio))}</td></tr>
  <tr><td>Conveyance Allowance</td><td>${fmt(sal.conveyance)}</td><td>${fmt(Math.round(sal.conveyance*ratio))}</td></tr>
  <tr><td>Medical Allowance</td><td>${fmt(sal.medical)}</td><td>${fmt(Math.round(sal.medical*ratio))}</td></tr>
  <tr><td>LTA</td><td>${fmt(sal.lta)}</td><td>${fmt(Math.round(sal.lta*ratio))}</td></tr>
  <tr><td>Other Allowance</td><td>${fmt(sal.other)}</td><td>${fmt(Math.round(sal.other*ratio))}</td></tr>
  <tr style="font-weight:700;background:#f8f9fa"><td>Gross Earnings</td><td>${fmt(sal.gross)}</td><td>${fmt(earnedGross)}</td></tr>
</table>
<div class="section-title">Deductions</div>
<table>
  <tr><th>Component</th><th>Amount</th></tr>
  <tr><td>Provident Fund (PF) — 12% of Basic</td><td>${fmt(pf)}</td></tr>
  <tr><td>ESI — 0.75% of Gross${sal.gross > 21000 ? ' (Not applicable — Gross > ₹21,000)' : ''}</td><td>${fmt(esi)}</td></tr>
  <tr><td>Professional Tax</td><td>${fmt(pt)}</td></tr>
  <tr><td>Labour Welfare Fund (LWF)</td><td>${fmt(lwf)}</td></tr>
  <tr style="font-weight:700;background:#f8f9fa"><td>Total Deductions</td><td>${fmt(totalDeductions)}</td></tr>
</table>
<table>
  <tr class="net-row"><td>NET PAY</td><td>${fmt(netPay)}</td></tr>
</table>
<p class="footer">Generated on ${_today()} · Workforce Operations Portal · This is a computer-generated payslip and does not require a signature.</p>
</body></html>`;

    return {
      success: true, html,
      summary: {
        gross: earnedGross, pf, esi, pt, lwf,
        totalDeductions, netPay,
        presentDays, workingDays, absentDays, lateDays, halfDays
      }
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// ─── BLOCK 2: OVERTIME & COMP-OFF TRACKING ──────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════
// Sheet: "Overtime" — OT_ID | EmpCode | EmpName | Date | ExtraHours | Reason | Status | AppliedDate | ApprovedBy | ApprovalDate | CompOffUsed
// Sheet: "Comp Off Balance" — EmpCode | Balance | LastUpdated

function submitOvertime(empCode, empName, date, extraHours, reason) {
  try {
    if (!date || !extraHours || Number(extraHours) <= 0) {
      return { success: false, message: 'Date and valid extra hours are required.' };
    }
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Overtime");
      const otId  = 'OT' + new Date().getTime();
      sheet.appendRow([otId, empCode, empName, date, Number(extraHours), reason || '', 'Pending', _today(), '', '', 'No']);
      _writeAuditLog(empCode, 'OT_SUBMITTED', 'Date: ' + date + ', Hours: ' + extraHours, '');
      return { success: true, message: 'Overtime request submitted.', otId };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

function getOvertimeRequests(empCode, role, filterStatus) {
  try {
    const data = _sheetData("Overtime");
    const records = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const rowEmp = String(data[i][1]);
      const rowSt  = String(data[i][6] || '');
      if (role === 'Employee' && rowEmp !== String(empCode)) continue;
      if (filterStatus && filterStatus !== 'All' && rowSt !== filterStatus) continue;
      records.push({
        otId:        String(data[i][0]),
        empCode:     String(data[i][1]),
        empName:     String(data[i][2]),
        date:        _formatDate(data[i][3]),
        extraHours:  Number(data[i][4]) || 0,
        reason:      String(data[i][5] || ''),
        status:      rowSt,
        appliedDate: _formatDate(data[i][7]),
        approvedBy:  String(data[i][8] || ''),
        compOffUsed: String(data[i][10] || 'No')
      });
    }
    records.sort((a, b) => b.otId.localeCompare(a.otId));
    return { success: true, records };
  } catch(e) { return { success: false, message: e.message }; }
}

function approveRejectOvertime(otId, action, approverName, approverCode) {
  try {
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Overtime");
      const data  = sheet.getDataRange().getValues();
      let targetRow = -1, row = null;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(otId)) { targetRow = i + 1; row = data[i]; break; }
      }
      if (targetRow === -1) return { success: false, message: 'OT request not found.' };
      if (String(row[6]) !== 'Pending') return { success: false, message: 'Already ' + row[6] + '.' };

      const newStatus = action === 'Approve' ? 'Approved' : 'Rejected';
      sheet.getRange(targetRow, 7).setValue(newStatus);
      sheet.getRange(targetRow, 9).setValue(approverName);
      sheet.getRange(targetRow, 10).setValue(_today());

      // Credit comp-off balance on approval (1 OT day = 1 comp-off day)
      if (newStatus === 'Approved') {
        const empCode = String(row[1]);
        const hours   = Number(row[4]) || 0;
        const compOff = hours >= 8 ? 1 : 0.5; // ≥8 hrs = 1 day, else 0.5
        try {
          const cbSheet = _getSheet("Comp Off Balance");
          const cbData  = cbSheet.getDataRange().getValues();
          let found = false;
          for (let i = 1; i < cbData.length; i++) {
            if (String(cbData[i][0]) === empCode) {
              const cur = Number(cbData[i][1]) || 0;
              cbSheet.getRange(i + 1, 2).setValue(cur + compOff);
              cbSheet.getRange(i + 1, 3).setValue(_today());
              found = true; break;
            }
          }
          if (!found) cbSheet.appendRow([empCode, compOff, _today()]);
        } catch(e) {}
      }

      _writeAuditLog(approverCode, 'OT_' + newStatus.toUpperCase(), 'OT_ID: ' + otId, '');
      return { success: true, message: 'Overtime ' + newStatus.toLowerCase() + '.' };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

function getCompOffBalance(empCode) {
  try {
    const data = _sheetData("Comp Off Balance");
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(empCode)) {
        return { success: true, balance: Number(data[i][1]) || 0, lastUpdated: _formatDate(data[i][2]) };
      }
    }
    return { success: true, balance: 0, lastUpdated: '' };
  } catch(e) { return { success: false, balance: 0, message: e.message }; }
}

// ═══════════════════════════════════════════════════════════════════════════
// ─── BLOCK 3: RESIGNATION / EXIT MANAGEMENT ─────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════
// Sheet: "Resignations" — ResID | EmpCode | EmpName | ResignDate | LastWorkingDay | Reason | Status | AppliedDate | HRNote | ProcessedBy | ProcessedDate
// Sheet: "Exit Checklist" — CheckID | EmpCode | Item | Status | CompletedBy | CompletedDate

function submitResignation(empCode, empName, lastWorkingDay, reason) {
  try {
    if (!lastWorkingDay || !reason || reason.trim().length < 5) {
      return { success: false, message: 'Last working day and reason (min 5 chars) are required.' };
    }
    const lock = _acquireLock();
    try {
      // Check for existing pending resignation
      const sheet = _getSheet("Resignations");
      const data  = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][1]) === String(empCode) && String(data[i][6]) === 'Pending') {
          return { success: false, message: 'You already have a pending resignation request.' };
        }
      }
      const resId = 'RES' + new Date().getTime();
      sheet.appendRow([resId, empCode, empName, _today(), lastWorkingDay, reason.trim(), 'Pending', _today(), '', '', '']);
      _writeAuditLog(empCode, 'RESIGNATION_SUBMITTED', 'LWD: ' + lastWorkingDay, '');
      return { success: true, message: 'Resignation submitted. HR will review and contact you.', resId };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

function getResignations(empCode, role) {
  try {
    const data = _sheetData("Resignations");
    const records = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (role === 'Employee' && String(data[i][1]) !== String(empCode)) continue;
      records.push({
        resId:          String(data[i][0]),
        empCode:        String(data[i][1]),
        empName:        String(data[i][2]),
        resignDate:     _formatDate(data[i][3]),
        lastWorkingDay: _formatDate(data[i][4]),
        reason:         String(data[i][5] || ''),
        status:         String(data[i][6] || 'Pending'),
        appliedDate:    _formatDate(data[i][7]),
        hrNote:         String(data[i][8] || ''),
        processedBy:    String(data[i][9] || ''),
        processedDate:  _formatDate(data[i][10])
      });
    }
    records.sort((a, b) => b.resId.localeCompare(a.resId));
    return { success: true, records };
  } catch(e) { return { success: false, message: e.message }; }
}

function processResignation(resId, action, hrNote, processorName, processorCode) {
  try {
    if (processorCode !== 'Admin' && processorCode !== 'Super Admin') {
      // role check via processorCode is actually the role string here
    }
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Resignations");
      const data  = sheet.getDataRange().getValues();
      let targetRow = -1, row = null;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(resId)) { targetRow = i + 1; row = data[i]; break; }
      }
      if (targetRow === -1) return { success: false, message: 'Resignation not found.' };

      const newStatus = action === 'Accept' ? 'Accepted' : 'Rejected';
      sheet.getRange(targetRow, 7).setValue(newStatus);
      sheet.getRange(targetRow, 9).setValue(hrNote || '');
      sheet.getRange(targetRow, 10).setValue(processorName);
      sheet.getRange(targetRow, 11).setValue(_today());

      // If accepted, create exit checklist items
      if (newStatus === 'Accepted') {
        const empCode = String(row[1]);
        const checkItems = ['ID Card Return', 'Laptop / Device Return', 'Access Card Return',
          'Pending Dues Cleared', 'Knowledge Transfer Completed', 'Exit Interview Done',
          'Final Settlement Initiated', 'Email Account Deactivation'];
        try {
          const clSheet = _getSheet("Exit Checklist");
          checkItems.forEach(item => {
            clSheet.appendRow(['CL' + new Date().getTime() + Math.random().toString(36).substr(2,4),
              empCode, item, 'Pending', '', '']);
          });
        } catch(e) {}
        // Update employee status to Resigned
        try { updateEmployeeStatus(empCode, 'Resigned', 'Super Admin'); } catch(e) {}
      }

      _writeAuditLog(processorName, 'RESIGNATION_' + newStatus.toUpperCase(), 'ResID: ' + resId, '');
      return { success: true, message: 'Resignation ' + newStatus.toLowerCase() + '.' };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

function getExitChecklist(empCode, requestorRole) {
  try {
    const data = _sheetData("Exit Checklist");
    const items = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (String(data[i][1]) !== String(empCode)) continue;
      items.push({
        checkId:       String(data[i][0]),
        empCode:       String(data[i][1]),
        item:          String(data[i][2]),
        status:        String(data[i][3] || 'Pending'),
        completedBy:   String(data[i][4] || ''),
        completedDate: _formatDate(data[i][5])
      });
    }
    return { success: true, items };
  } catch(e) { return { success: false, message: e.message }; }
}

function updateExitChecklistItem(checkId, status, completedBy) {
  try {
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Exit Checklist");
      const data  = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(checkId)) {
          sheet.getRange(i + 1, 4).setValue(status);
          sheet.getRange(i + 1, 5).setValue(completedBy);
          sheet.getRange(i + 1, 6).setValue(_today());
          return { success: true, message: 'Checklist item updated.' };
        }
      }
      return { success: false, message: 'Item not found.' };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

// ═══════════════════════════════════════════════════════════════════════════
// ─── BLOCK 4: PERFORMANCE MANAGEMENT (KRA/KPI) ──────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════
// Sheet: "Performance" — PerfID | EmpCode | EmpName | Period | KRA | Target | SelfRating | ManagerRating | FinalGrade | Comments | Status | CreatedBy | CreatedDate

function addPerformanceGoal(data, creatorRole) {
  try {
    if (creatorRole !== 'Manager' && creatorRole !== 'Admin' && creatorRole !== 'Super Admin') {
      return { success: false, message: 'Access denied.' };
    }
    const lock = _acquireLock();
    try {
      const sheet  = _getSheet("Performance");
      const perfId = 'PERF' + new Date().getTime();
      sheet.appendRow([
        perfId, data.empCode, data.empName, data.period,
        data.kra, data.target, '', '', '', data.comments || '',
        'Active', data.createdBy, _today()
      ]);
      _writeAuditLog(data.createdBy, 'PERF_GOAL_ADDED', 'EmpCode: ' + data.empCode + ', KRA: ' + data.kra, '');
      return { success: true, message: 'Performance goal added.', perfId };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

function getPerformanceGoals(empCode, role, period) {
  try {
    const data = _sheetData("Performance");
    const records = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (role === 'Employee' && String(data[i][1]) !== String(empCode)) continue;
      if (period && period !== 'All' && String(data[i][3]) !== period) continue;
      records.push({
        perfId:        String(data[i][0]),
        empCode:       String(data[i][1]),
        empName:       String(data[i][2]),
        period:        String(data[i][3]),
        kra:           String(data[i][4]),
        target:        String(data[i][5]),
        selfRating:    String(data[i][6] || ''),
        managerRating: String(data[i][7] || ''),
        finalGrade:    String(data[i][8] || ''),
        comments:      String(data[i][9] || ''),
        status:        String(data[i][10] || 'Active'),
        createdBy:     String(data[i][11] || ''),
        createdDate:   _formatDate(data[i][12])
      });
    }
    records.sort((a, b) => b.perfId.localeCompare(a.perfId));
    return { success: true, records };
  } catch(e) { return { success: false, message: e.message }; }
}

function submitSelfAppraisal(perfId, selfRating, comments, empCode) {
  try {
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Performance");
      const data  = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(perfId)) {
          if (String(data[i][1]) !== String(empCode)) return { success: false, message: 'Access denied.' };
          sheet.getRange(i + 1, 7).setValue(selfRating);
          if (comments) sheet.getRange(i + 1, 10).setValue(comments);
          _writeAuditLog(empCode, 'SELF_APPRAISAL', 'PerfID: ' + perfId + ', Rating: ' + selfRating, '');
          return { success: true, message: 'Self-appraisal submitted.' };
        }
      }
      return { success: false, message: 'Goal not found.' };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

function submitManagerRating(perfId, managerRating, finalGrade, comments, managerCode) {
  try {
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Performance");
      const data  = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(perfId)) {
          sheet.getRange(i + 1, 8).setValue(managerRating);
          sheet.getRange(i + 1, 9).setValue(finalGrade);
          if (comments) sheet.getRange(i + 1, 10).setValue(comments);
          sheet.getRange(i + 1, 11).setValue('Reviewed');
          _writeAuditLog(managerCode, 'MANAGER_RATING', 'PerfID: ' + perfId + ', Grade: ' + finalGrade, '');
          return { success: true, message: 'Manager rating submitted.' };
        }
      }
      return { success: false, message: 'Goal not found.' };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

// ═══════════════════════════════════════════════════════════════════════════
// ─── BLOCK 5: TRAINING & CERTIFICATION TRACKER ──────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════
// Sheet: "Training" — TrainID | EmpCode | EmpName | TrainingName | Provider | StartDate | EndDate | CertExpiry | Status | Score | Notes | AddedBy | AddedDate

function addTrainingRecord(data, adderRole) {
  try {
    const lock = _acquireLock();
    try {
      const sheet   = _getSheet("Training");
      const trainId = 'TRN' + new Date().getTime();
      sheet.appendRow([
        trainId, data.empCode, data.empName, data.trainingName,
        data.provider || '', data.startDate || '', data.endDate || '',
        data.certExpiry || '', data.status || 'Completed',
        data.score || '', data.notes || '', data.addedBy, _today()
      ]);
      _writeAuditLog(data.addedBy, 'TRAINING_ADDED', 'EmpCode: ' + data.empCode + ', Training: ' + data.trainingName, '');
      return { success: true, message: 'Training record added.', trainId };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

function getTrainingRecords(empCode, role) {
  try {
    const data = _sheetData("Training");
    const records = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (role === 'Employee' && String(data[i][1]) !== String(empCode)) continue;
      const certExpiry = _formatDate(data[i][7]);
      // Flag expiring within 30 days
      let expiryAlert = false;
      if (certExpiry) {
        const parts = certExpiry.split('/');
        if (parts.length === 3) {
          const expDate = new Date(Number(parts[2]), Number(parts[1])-1, Number(parts[0]));
          const diff = (expDate - new Date()) / (1000*60*60*24);
          if (diff >= 0 && diff <= 30) expiryAlert = true;
        }
      }
      records.push({
        trainId:      String(data[i][0]),
        empCode:      String(data[i][1]),
        empName:      String(data[i][2]),
        trainingName: String(data[i][3]),
        provider:     String(data[i][4] || ''),
        startDate:    _formatDate(data[i][5]),
        endDate:      _formatDate(data[i][6]),
        certExpiry,
        status:       String(data[i][8] || 'Completed'),
        score:        String(data[i][9] || ''),
        notes:        String(data[i][10] || ''),
        addedBy:      String(data[i][11] || ''),
        addedDate:    _formatDate(data[i][12]),
        expiryAlert
      });
    }
    records.sort((a, b) => b.trainId.localeCompare(a.trainId));
    return { success: true, records };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteTrainingRecord(trainId, role) {
  try {
    if (role !== 'Admin' && role !== 'Super Admin' && role !== 'Manager') {
      return { success: false, message: 'Access denied.' };
    }
    const sheet = _getSheet("Training");
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(trainId)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Training record deleted.' };
      }
    }
    return { success: false, message: 'Record not found.' };
  } catch(e) { return { success: false, message: e.message }; }
}

// ═══════════════════════════════════════════════════════════════════════════
// ─── BLOCK 6: ASSET MANAGEMENT ──────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════
// Sheet: "Assets" — AssetID | AssetName | AssetType | SerialNo | AssignedTo | AssignedDate | Condition | Status | ReturnDate | Notes

function addAsset(data, adderRole) {
  try {
    if (adderRole !== 'Admin' && adderRole !== 'Super Admin') {
      return { success: false, message: 'Access denied.' };
    }
    const lock = _acquireLock();
    try {
      const sheet   = _getSheet("Assets");
      const assetId = 'AST' + new Date().getTime();
      sheet.appendRow([
        assetId, data.assetName, data.assetType || 'Other',
        data.serialNo || '', data.assignedTo || '',
        data.assignedTo ? _today() : '',
        data.condition || 'Good',
        data.assignedTo ? 'Assigned' : 'Available',
        '', data.notes || ''
      ]);
      _writeAuditLog(data.addedBy || adderRole, 'ASSET_ADDED', 'Asset: ' + data.assetName, '');
      return { success: true, message: 'Asset added.', assetId };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

function getAssets(empCode, role) {
  try {
    const data = _sheetData("Assets");
    const assets = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      // Employees see only their assigned assets
      if (role === 'Employee' && String(data[i][4]) !== String(empCode)) continue;
      assets.push({
        assetId:      String(data[i][0]),
        assetName:    String(data[i][1]),
        assetType:    String(data[i][2] || ''),
        serialNo:     String(data[i][3] || ''),
        assignedTo:   String(data[i][4] || ''),
        assignedDate: _formatDate(data[i][5]),
        condition:    String(data[i][6] || 'Good'),
        status:       String(data[i][7] || 'Available'),
        returnDate:   _formatDate(data[i][8]),
        notes:        String(data[i][9] || '')
      });
    }
    return { success: true, assets };
  } catch(e) { return { success: false, message: e.message }; }
}

function assignAsset(assetId, empCode, adderRole) {
  try {
    if (adderRole !== 'Admin' && adderRole !== 'Super Admin') {
      return { success: false, message: 'Access denied.' };
    }
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Assets");
      const data  = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(assetId)) {
          sheet.getRange(i + 1, 5).setValue(empCode);
          sheet.getRange(i + 1, 6).setValue(_today());
          sheet.getRange(i + 1, 8).setValue('Assigned');
          sheet.getRange(i + 1, 9).setValue('');
          _writeAuditLog(adderRole, 'ASSET_ASSIGNED', 'AssetID: ' + assetId + ' → ' + empCode, '');
          return { success: true, message: 'Asset assigned.' };
        }
      }
      return { success: false, message: 'Asset not found.' };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

function returnAsset(assetId, condition, adderRole) {
  try {
    if (adderRole !== 'Admin' && adderRole !== 'Super Admin') {
      return { success: false, message: 'Access denied.' };
    }
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Assets");
      const data  = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(assetId)) {
          sheet.getRange(i + 1, 5).setValue('');
          sheet.getRange(i + 1, 7).setValue(condition || 'Good');
          sheet.getRange(i + 1, 8).setValue('Available');
          sheet.getRange(i + 1, 9).setValue(_today());
          _writeAuditLog(adderRole, 'ASSET_RETURNED', 'AssetID: ' + assetId, '');
          return { success: true, message: 'Asset marked as returned.' };
        }
      }
      return { success: false, message: 'Asset not found.' };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

// ═══════════════════════════════════════════════════════════════════════════
// ─── BLOCK 7: EXPENSE REIMBURSEMENT ─────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════
// Sheet: "Expenses" — ExpID | EmpCode | EmpName | Category | Amount | Date | Description | ReceiptDriveId | Status | AppliedDate | ManagerApproval | ManagerDate | FinanceApproval | FinanceDate | Notes

function submitExpense(data) {
  try {
    if (!data.amount || Number(data.amount) <= 0) {
      return { success: false, message: 'Valid amount is required.' };
    }
    const lock = _acquireLock();
    try {
      const sheet  = _getSheet("Expenses");
      const expId  = 'EXP' + new Date().getTime();
      let receiptId = '';

      // Upload receipt to Drive if provided
      if (data.receiptBase64 && data.receiptMime) {
        try {
          const blob   = Utilities.newBlob(Utilities.base64Decode(data.receiptBase64), data.receiptMime, data.receiptFileName || 'receipt');
          const folder = DriveApp.getFoldersByName('HRMS_Expenses').hasNext()
            ? DriveApp.getFoldersByName('HRMS_Expenses').next()
            : DriveApp.createFolder('HRMS_Expenses');
          const file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          receiptId = file.getId();
        } catch(e) {}
      }

      sheet.appendRow([
        expId, data.empCode, data.empName, data.category || 'Other',
        Number(data.amount), data.date || _today(), data.description || '',
        receiptId, 'Pending', _today(), '', '', '', '', ''
      ]);
      _writeAuditLog(data.empCode, 'EXPENSE_SUBMITTED', 'Amount: ' + data.amount + ', Category: ' + data.category, '');
      return { success: true, message: 'Expense claim submitted.', expId };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

function getExpenses(empCode, role, filterStatus) {
  try {
    const data = _sheetData("Expenses");
    const records = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (role === 'Employee' && String(data[i][1]) !== String(empCode)) continue;
      const rowSt = String(data[i][8] || 'Pending');
      if (filterStatus && filterStatus !== 'All' && rowSt !== filterStatus) continue;
      const receiptId = String(data[i][7] || '');
      records.push({
        expId:           String(data[i][0]),
        empCode:         String(data[i][1]),
        empName:         String(data[i][2]),
        category:        String(data[i][3] || ''),
        amount:          Number(data[i][4]) || 0,
        date:            _formatDate(data[i][5]),
        description:     String(data[i][6] || ''),
        receiptUrl:      receiptId ? 'https://drive.google.com/uc?export=view&id=' + receiptId : '',
        status:          rowSt,
        appliedDate:     _formatDate(data[i][9]),
        managerApproval: String(data[i][10] || ''),
        financeApproval: String(data[i][12] || ''),
        notes:           String(data[i][14] || '')
      });
    }
    records.sort((a, b) => b.expId.localeCompare(a.expId));
    return { success: true, records };
  } catch(e) { return { success: false, message: e.message }; }
}

function approveExpense(expId, stage, action, approverName, approverCode) {
  // stage: 'Manager' or 'Finance'
  try {
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Expenses");
      const data  = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) !== String(expId)) continue;
        const newStatus = action === 'Approve' ? 'Approved' : 'Rejected';
        if (stage === 'Manager') {
          sheet.getRange(i + 1, 11).setValue(approverName + ' — ' + newStatus);
          sheet.getRange(i + 1, 12).setValue(_today());
          sheet.getRange(i + 1, 9).setValue(action === 'Approve' ? 'Manager Approved' : 'Rejected');
        } else {
          sheet.getRange(i + 1, 13).setValue(approverName + ' — ' + newStatus);
          sheet.getRange(i + 1, 14).setValue(_today());
          sheet.getRange(i + 1, 9).setValue(action === 'Approve' ? 'Finance Approved' : 'Rejected');
        }
        _writeAuditLog(approverCode, 'EXPENSE_' + stage.toUpperCase() + '_' + newStatus.toUpperCase(), 'ExpID: ' + expId, '');
        return { success: true, message: 'Expense ' + newStatus.toLowerCase() + ' by ' + stage + '.' };
      }
      return { success: false, message: 'Expense not found.' };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

// ═══════════════════════════════════════════════════════════════════════════
// ─── BLOCK 8: ANALYTICS — HEADCOUNT, LEAVE UTILIZATION, BIRTHDAY ────────────
// ═══════════════════════════════════════════════════════════════════════════

function getHeadcountReport(requestorRole) {
  try {
    if (requestorRole !== 'Admin' && requestorRole !== 'Super Admin' && requestorRole !== 'Manager') {
      return { success: false, message: 'Access denied.' };
    }
    const empData = _sheetData("Employee Master");
    const h = empData[0];
    const idCol     = _colIdx(h, "Emp ID");
    const nameCol   = _colIdx(h, "Employee Name");
    const levelCol  = _colIdx(h, "Level");
    const statusCol = _colIdx(h, "Status");
    const genderCol = _colIdx(h, "Gender");
    const dojCol    = _colIdx(h, "Date of Joining");

    let total = 0, active = 0, inactive = 0, resigned = 0;
    const deptMap = {}, genderMap = { Male: 0, Female: 0, Other: 0 };
    const tenureBuckets = { '< 1 yr': 0, '1–3 yrs': 0, '3–5 yrs': 0, '5+ yrs': 0 };
    const now = new Date();

    for (let i = 1; i < empData.length; i++) {
      if (!empData[i][idCol]) continue;
      total++;
      const st = String(empData[i][statusCol] || '').toLowerCase();
      if (st === 'active') active++;
      else if (st === 'resigned') resigned++;
      else inactive++;

      const dept = String(empData[i][levelCol] || 'Unknown');
      deptMap[dept] = (deptMap[dept] || 0) + 1;

      const gender = String(empData[i][genderCol] || '').trim();
      if (gender === 'Male') genderMap.Male++;
      else if (gender === 'Female') genderMap.Female++;
      else genderMap.Other++;

      // Tenure
      const doj = dojCol >= 0 ? empData[i][dojCol] : null;
      if (doj) {
        let dojDate;
        if (doj instanceof Date) dojDate = doj;
        else {
          const parts = String(doj).split('/');
          if (parts.length === 3) dojDate = new Date(Number(parts[2]), Number(parts[1])-1, Number(parts[0]));
        }
        if (dojDate) {
          const yrs = (now - dojDate) / (1000*60*60*24*365);
          if (yrs < 1) tenureBuckets['< 1 yr']++;
          else if (yrs < 3) tenureBuckets['1–3 yrs']++;
          else if (yrs < 5) tenureBuckets['3–5 yrs']++;
          else tenureBuckets['5+ yrs']++;
        }
      }
    }

    const deptBreakdown = Object.keys(deptMap).map(k => ({ dept: k, count: deptMap[k] }))
      .sort((a, b) => b.count - a.count);

    return {
      success: true,
      total, active, inactive, resigned,
      deptBreakdown, genderMap, tenureBuckets
    };
  } catch(e) { return { success: false, message: e.message }; }
}

function getLeaveUtilizationReport(requestorRole) {
  try {
    if (requestorRole !== 'Admin' && requestorRole !== 'Super Admin' && requestorRole !== 'Manager') {
      return { success: false, message: 'Access denied.' };
    }
    const lbData  = _sheetData("Leave Balance");
    const empData = _sheetData("Employee Master");
    const h = empData[0];
    const idCol   = _colIdx(h, "Emp ID");
    const nameCol = _colIdx(h, "Employee Name");
    const empMap  = {};
    for (let i = 1; i < empData.length; i++) {
      if (empData[i][idCol]) empMap[String(empData[i][idCol])] = String(empData[i][nameCol] || '');
    }

    const rows = [];
    for (let i = 1; i < lbData.length; i++) {
      if (!lbData[i][0]) continue;
      const empCode = String(lbData[i][0]);
      const cl = Number(lbData[i][1]) || 0;
      const sl = Number(lbData[i][2]) || 0;
      const pl = Number(lbData[i][3]) || 0;
      const total = cl + sl + pl;
      rows.push({
        empCode,
        name:    empMap[empCode] || empCode,
        cl, sl, pl, total,
        // Flag high balance (potential encashment eligibility)
        highBalance: total > 20
      });
    }
    rows.sort((a, b) => b.total - a.total);
    return { success: true, rows };
  } catch(e) { return { success: false, message: e.message }; }
}

function getUpcomingBirthdays(daysAhead) {
  try {
    const empData = _sheetData("Employee Master");
    const h = empData[0];
    const idCol     = _colIdx(h, "Emp ID");
    const nameCol   = _colIdx(h, "Employee Name");
    const dobCol    = _colIdx(h, "DOB");
    const dojCol    = _colIdx(h, "Date of Joining");
    const statusCol = _colIdx(h, "Status");

    const ahead = daysAhead || 30;
    const now   = new Date();
    const results = { birthdays: [], anniversaries: [] };

    for (let i = 1; i < empData.length; i++) {
      if (!empData[i][idCol]) continue;
      const st = String(empData[i][statusCol] || '').toLowerCase();
      if (st !== 'active') continue;

      const empCode = String(empData[i][idCol]);
      const name    = String(empData[i][nameCol] || '');

      // Birthday check
      if (dobCol >= 0 && empData[i][dobCol]) {
        let dob;
        const raw = empData[i][dobCol];
        if (raw instanceof Date) dob = raw;
        else {
          const parts = String(raw).split('/');
          if (parts.length === 3) dob = new Date(Number(parts[2]), Number(parts[1])-1, Number(parts[0]));
        }
        if (dob) {
          const thisYearBday = new Date(now.getFullYear(), dob.getMonth(), dob.getDate());
          if (thisYearBday < now) thisYearBday.setFullYear(now.getFullYear() + 1);
          const diff = Math.round((thisYearBday - now) / (1000*60*60*24));
          if (diff >= 0 && diff <= ahead) {
            results.birthdays.push({ empCode, name, date: _formatDate(dob), daysAway: diff });
          }
        }
      }

      // Work anniversary check
      if (dojCol >= 0 && empData[i][dojCol]) {
        let doj;
        const raw = empData[i][dojCol];
        if (raw instanceof Date) doj = raw;
        else {
          const parts = String(raw).split('/');
          if (parts.length === 3) doj = new Date(Number(parts[2]), Number(parts[1])-1, Number(parts[0]));
        }
        if (doj) {
          const thisYearAnn = new Date(now.getFullYear(), doj.getMonth(), doj.getDate());
          if (thisYearAnn < now) thisYearAnn.setFullYear(now.getFullYear() + 1);
          const diff = Math.round((thisYearAnn - now) / (1000*60*60*24));
          const years = now.getFullYear() - doj.getFullYear() + (thisYearAnn.getFullYear() > now.getFullYear() ? 0 : 0);
          if (diff >= 0 && diff <= ahead) {
            results.anniversaries.push({ empCode, name, date: _formatDate(doj), daysAway: diff, years: Math.max(1, now.getFullYear() - doj.getFullYear()) });
          }
        }
      }
    }

    results.birthdays.sort((a, b) => a.daysAway - b.daysAway);
    results.anniversaries.sort((a, b) => a.daysAway - b.daysAway);
    return { success: true, ...results };
  } catch(e) { return { success: false, message: e.message }; }
}

// ═══════════════════════════════════════════════════════════════════════════
// ─── BLOCK 9: EMERGENCY CONTACT, GRIEVANCE, ROLE ASSIGNMENT, BULK OPS ───────
// ═══════════════════════════════════════════════════════════════════════════

// ── Emergency Contact ────────────────────────────────────────────────────────
// Sheet: "Emergency Contacts" — EmpCode | ContactName | Relationship | Phone | AltPhone | Address

function saveEmergencyContact(empCode, data) {
  try {
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Emergency Contacts");
      const rows  = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(empCode)) {
          sheet.getRange(i + 1, 2, 1, 5).setValues([[
            data.contactName, data.relationship, data.phone, data.altPhone || '', data.address || ''
          ]]);
          _writeAuditLog(empCode, 'EMERGENCY_CONTACT_UPDATED', '', '');
          return { success: true, message: 'Emergency contact updated.' };
        }
      }
      sheet.appendRow([empCode, data.contactName, data.relationship, data.phone, data.altPhone || '', data.address || '']);
      _writeAuditLog(empCode, 'EMERGENCY_CONTACT_ADDED', '', '');
      return { success: true, message: 'Emergency contact saved.' };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

function getEmergencyContact(empCode, requestorCode, requestorRole) {
  try {
    if (requestorRole === 'Employee' && requestorCode !== empCode) {
      return { success: false, message: 'Access denied.' };
    }
    const data = _sheetData("Emergency Contacts");
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(empCode)) {
        return {
          success: true,
          contact: {
            contactName:  String(data[i][1] || ''),
            relationship: String(data[i][2] || ''),
            phone:        String(data[i][3] || ''),
            altPhone:     String(data[i][4] || ''),
            address:      String(data[i][5] || '')
          }
        };
      }
    }
    return { success: true, contact: null };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── Grievance System ─────────────────────────────────────────────────────────
// Sheet: "Grievances" — GrievID | EmpCode | EmpName | Category | Description | Anonymous | Status | SubmittedDate | HRNote | ResolvedBy | ResolvedDate

function submitGrievance(empCode, empName, category, description, isAnonymous) {
  try {
    if (!description || description.trim().length < 10) {
      return { success: false, message: 'Description must be at least 10 characters.' };
    }
    const lock = _acquireLock();
    try {
      const sheet    = _getSheet("Grievances");
      const grievId  = 'GRV' + new Date().getTime();
      const displayName = isAnonymous ? 'Anonymous' : empName;
      const displayCode = isAnonymous ? 'ANON' : empCode;
      sheet.appendRow([grievId, displayCode, displayName, category || 'General', description.trim(), isAnonymous ? 'Yes' : 'No', 'Open', _today(), '', '', '']);
      _writeAuditLog(isAnonymous ? 'ANONYMOUS' : empCode, 'GRIEVANCE_SUBMITTED', 'Category: ' + category, '');
      return { success: true, message: 'Grievance submitted. HR will review within 5 working days.', grievId };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

function getGrievances(empCode, role) {
  try {
    const data = _sheetData("Grievances");
    const records = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      // Employees see only their own (non-anonymous) grievances
      if (role === 'Employee' && String(data[i][1]) !== String(empCode)) continue;
      records.push({
        grievId:       String(data[i][0]),
        empCode:       String(data[i][1]),
        empName:       String(data[i][2]),
        category:      String(data[i][3] || ''),
        description:   String(data[i][4] || ''),
        anonymous:     String(data[i][5] || 'No'),
        status:        String(data[i][6] || 'Open'),
        submittedDate: _formatDate(data[i][7]),
        hrNote:        String(data[i][8] || ''),
        resolvedBy:    String(data[i][9] || ''),
        resolvedDate:  _formatDate(data[i][10])
      });
    }
    records.sort((a, b) => b.grievId.localeCompare(a.grievId));
    return { success: true, records };
  } catch(e) { return { success: false, message: e.message }; }
}

function resolveGrievance(grievId, hrNote, resolverName, resolverCode) {
  try {
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Grievances");
      const data  = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(grievId)) {
          sheet.getRange(i + 1, 7).setValue('Resolved');
          sheet.getRange(i + 1, 9).setValue(hrNote || '');
          sheet.getRange(i + 1, 10).setValue(resolverName);
          sheet.getRange(i + 1, 11).setValue(_today());
          _writeAuditLog(resolverCode, 'GRIEVANCE_RESOLVED', 'GrievID: ' + grievId, '');
          return { success: true, message: 'Grievance marked as resolved.' };
        }
      }
      return { success: false, message: 'Grievance not found.' };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

// ── Role Assignment UI ────────────────────────────────────────────────────────
// Sheet: "Role" — RoleID | Role | EmpCode

function getAllRoles(requestorRole) {
  try {
    if (requestorRole !== 'Super Admin') return { success: false, message: 'Access denied.' };
    const data = _sheetData("Role");
    const roles = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      roles.push({ roleId: String(data[i][0]), role: String(data[i][1]), empCode: String(data[i][2]) });
    }
    return { success: true, roles };
  } catch(e) { return { success: false, message: e.message }; }
}

function assignRole(empCode, role, assignerRole) {
  try {
    if (assignerRole !== 'Super Admin') return { success: false, message: 'Only Super Admin can assign roles.' };
    const validRoles = ['Employee', 'Manager', 'Admin', 'Super Admin'];
    if (!validRoles.includes(role)) return { success: false, message: 'Invalid role.' };

    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Role");
      const data  = sheet.getDataRange().getValues();
      // Update existing row if found
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][2]) === String(empCode)) {
          sheet.getRange(i + 1, 2).setValue(role);
          _invalidateCache("Role");
          _writeAuditLog(assignerRole, 'ROLE_ASSIGNED', 'EmpCode: ' + empCode + ' → ' + role, '');
          return { success: true, message: 'Role updated to ' + role + '.' };
        }
      }
      // Add new row
      const roleId = 'ROLE' + new Date().getTime();
      sheet.appendRow([roleId, role, empCode]);
      _invalidateCache("Role");
      _writeAuditLog(assignerRole, 'ROLE_ASSIGNED', 'EmpCode: ' + empCode + ' → ' + role, '');
      return { success: true, message: 'Role ' + role + ' assigned.' };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

function removeRole(empCode, assignerRole) {
  try {
    if (assignerRole !== 'Super Admin') return { success: false, message: 'Access denied.' };
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Role");
      const data  = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][2]) === String(empCode)) {
          sheet.deleteRow(i + 1);
          _invalidateCache("Role");
          _writeAuditLog(assignerRole, 'ROLE_REMOVED', 'EmpCode: ' + empCode, '');
          return { success: true, message: 'Role removed. Employee reverts to default Employee role.' };
        }
      }
      return { success: false, message: 'No role record found for this employee.' };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

// ── Bulk Leave Balance Update ─────────────────────────────────────────────────
function bulkUpdateLeaveBalance(updates, requestorRole) {
  // updates: [{ empCode, cl, sl, pl }]
  try {
    if (requestorRole !== 'Admin' && requestorRole !== 'Super Admin') {
      return { success: false, message: 'Access denied.' };
    }
    const lock = _acquireLock();
    try {
      const sheet = _getSheet("Leave Balance");
      const data  = sheet.getDataRange().getValues();
      let updated = 0;
      updates.forEach(upd => {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][0]) === String(upd.empCode)) {
            if (upd.cl !== undefined) sheet.getRange(i + 1, 2).setValue(Number(upd.cl));
            if (upd.sl !== undefined) sheet.getRange(i + 1, 3).setValue(Number(upd.sl));
            if (upd.pl !== undefined) sheet.getRange(i + 1, 4).setValue(Number(upd.pl));
            sheet.getRange(i + 1, 5).setValue(_today());
            updated++;
            break;
          }
        }
      });
      _writeAuditLog(requestorRole, 'BULK_LEAVE_BALANCE_UPDATE', 'Updated: ' + updated + ' records', '');
      return { success: true, message: updated + ' leave balance records updated.' };
    } finally { _releaseLock(lock); }
  } catch(e) { return { success: false, message: e.message }; }
}

// ── Email Preferences (stub — stores in Script Properties per employee) ───────
function getEmailPrefs(empCode) {
  try {
    const props = PropertiesService.getUserProperties();
    const raw   = props.getProperty('email_prefs_' + empCode);
    if (raw) return JSON.parse(raw);
    return { enabled: true, leave: true, gatepass: true, regularization: true, tasks: true };
  } catch(e) {
    return { enabled: true, leave: true, gatepass: true, regularization: true, tasks: true };
  }
}

function saveEmailPrefs(empCode, prefs) {
  try {
    PropertiesService.getUserProperties().setProperty('email_prefs_' + empCode, JSON.stringify(prefs));
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── Scheduled Trigger: Auto-mark absent + Birthday announcements ──────────────
/**
 * Run this daily via a time-based trigger:
 *   Triggers → Add Trigger → dailyScheduledTasks → Time-driven → Day timer → 11 PM
 */
function dailyScheduledTasks() {
  try {
    _autoMarkAbsent();
    _postBirthdayAnniversaryAnnouncements();
  } catch(e) {
    Logger.log('dailyScheduledTasks error: ' + e.message);
  }
}

function _autoMarkAbsent() {
  try {
    const today    = _today();
    const empData  = _sheetData("Employee Master");
    const h        = empData[0];
    const idCol    = _colIdx(h, "Emp ID");
    const statCol  = _colIdx(h, "Status");
    const attSheet = _getSheet("Attendance Sheet");
    const attData  = attSheet.getDataRange().getValues();

    // Build set of employees who already have an attendance record today
    const checkedIn = new Set();
    for (let i = 1; i < attData.length; i++) {
      if (_formatDate(attData[i][2]) === today) checkedIn.add(String(attData[i][1]));
    }

    // For each active employee without a record today, mark absent
    for (let i = 1; i < empData.length; i++) {
      if (!empData[i][idCol]) continue;
      const st = String(empData[i][statCol] || '').toLowerCase();
      if (st !== 'active') continue;
      const empCode = String(empData[i][idCol]);
      if (checkedIn.has(empCode)) continue;

      // Check if today is a holiday or Sunday
      const now = new Date();
      if (now.getDay() === 0) continue; // Sunday
      const holData = _sheetData("Holiday List");
      let isHoliday = false;
      for (let j = 1; j < holData.length; j++) {
        if (_formatDate(holData[j][1]) === today) { isHoliday = true; break; }
      }
      if (isHoliday) continue;

      const attId = 'ATT' + new Date().getTime() + i;
      attSheet.appendRow([attId, empCode, today, '', '', '', 'Auto-Absent', 'Absent', '', '']);
    }
  } catch(e) { Logger.log('_autoMarkAbsent error: ' + e.message); }
}

function _postBirthdayAnniversaryAnnouncements() {
  try {
    const result = getUpcomingBirthdays(0); // today only
    if (!result.success) return;

    const annSheet = _getSpreadsheet().getSheetByName("Announcements") ||
      (() => {
        const s = _getSpreadsheet().insertSheet("Announcements");
        s.appendRow(['AnnID','Title','Content','Category','PostedBy','PostedDate','ExpiryDate','Priority']);
        return s;
      })();

    result.birthdays.forEach(b => {
      if (b.daysAway === 0) {
        const annId = 'ANN' + new Date().getTime();
        annSheet.appendRow([annId, '🎂 Happy Birthday ' + b.name + '!',
          'Wishing ' + b.name + ' (' + b.empCode + ') a very Happy Birthday! 🎉',
          'HR', 'System', _today(), _today(), 'Normal']);
      }
    });

    result.anniversaries.forEach(a => {
      if (a.daysAway === 0) {
        const annId = 'ANN' + new Date().getTime();
        annSheet.appendRow([annId, '🏆 Work Anniversary — ' + a.name,
          'Congratulations to ' + a.name + ' on completing ' + a.years + ' year(s) with us! 🎊',
          'HR', 'System', _today(), _today(), 'Normal']);
      }
    });
    _invalidateCache("Announcements");
  } catch(e) { Logger.log('_postBirthdayAnniversaryAnnouncements error: ' + e.message); }
}
