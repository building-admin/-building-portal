// ════════════════════════════════════════════════════
//  Roster.gs — CRUD ตารางรอบและพนักงาน
// ════════════════════════════════════════════════════

// ════════════════════════════════════════════════════
//  GET ROSTER — ดึงตารางรอบตามเดือน
// ════════════════════════════════════════════════════
function getRoster(month) {
  try {
    if (!month) return { success: false, message: 'ไม่ระบุเดือน' };

    const sh   = getSheet(SHEET_ROSTER);
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };

    const headers = data[0].map(h => String(h).trim());

    // แปลง month format ให้รองรับทั้ง "2569-05" และ Date object จาก Sheets
    function normalizeMonth(val) {
      if (!val) return '';
      if (val instanceof Date) {
        const y = val.getFullYear() + 543;
        const m = String(val.getMonth() + 1).padStart(2, '0');
        return `${y}-${m}`;
      }
      const s = String(val).trim();
      // รูปแบบ d/m/yyyy หรือ d/m/yyyy (พ.ศ.)
      const match = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (match) {
        const m = String(match[2]).padStart(2, '0');
        return `${match[3]}-${m}`;
      }
      return s;
    }

    const rows = data.slice(1)
      .filter(row => normalizeMonth(row[0]) === String(month).trim())
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = row[i]; });
        return obj;
      });

    return { success: true, data: rows, month };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ════════════════════════════════════════════════════
//  SAVE ROSTER — บันทึกตารางรอบทั้งเดือน
//  data = array of { รหัสพนักงาน, ชื่อ, ตำแหน่ง, วันหยุดประจำ, 1..31 }
// ════════════════════════════════════════════════════
function saveRoster(month, data) {
  try {
    if (!month || !data) return { success: false, message: 'ข้อมูลไม่ครบ' };

    const sh      = getSheet(SHEET_ROSTER);
    const allData = sh.getDataRange().getValues();
    const headers = allData[0].map(h => String(h).trim());

    // ลบแถวเดือนนี้ออกก่อน (จากล่างขึ้นบน)
    for (let i = allData.length - 1; i >= 1; i--) {
      if (String(allData[i][0]).trim() === String(month).trim()) {
        sh.deleteRow(i + 1);
      }
    }

    // เขียนข้อมูลใหม่
    data.forEach(emp => {
      const row = headers.map(h => {
        if (h === 'เดือน') return String(month); // เก็บเป็น string เสมอ
        return emp[h] !== undefined ? emp[h] : '';
      });
      sh.appendRow(row);
    });

    return { success: true, message: `บันทึกตารางเดือน ${month} เรียบร้อยแล้ว` };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ════════════════════════════════════════════════════
//  GET EMPLOYEES — ดึงรายชื่อพนักงาน active
// ════════════════════════════════════════════════════
function getEmployees() {
  try {
    const emps = getSheetData(SHEET_EMP)
      .filter(e => String(e['สถานะ']).toLowerCase() === 'active')
      .sort((a, b) => parseInt(a['ลำดับ'] || 0) - parseInt(b['ลำดับ'] || 0));
    return { success: true, data: emps };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ════════════════════════════════════════════════════
//  ADD EMPLOYEE
// ════════════════════════════════════════════════════
function addEmployee(emp) {
  try {
    if (!emp || !emp['รหัสพนักงาน']) return { success: false, message: 'ข้อมูลไม่ครบ' };

    // ตรวจสอบซ้ำ
    const existing = getSheetData(SHEET_EMP);
    if (existing.find(e => String(e['รหัสพนักงาน']) === String(emp['รหัสพนักงาน']))) {
      return { success: false, message: 'รหัสพนักงานนี้มีอยู่แล้ว' };
    }

    const sh = getSheet(SHEET_EMP);
    sh.appendRow([
      emp['รหัสพนักงาน'],
      emp['ชื่อ-นามสกุล']  || '',
      emp['ตำแหน่ง']       || '',
      emp['แผนก']          || '',
      emp['วันหยุดประจำ']  || '',
      emp['ลำดับ']         || existing.length + 1,
      'active'
    ]);

    return { success: true, message: 'เพิ่มพนักงานเรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ════════════════════════════════════════════════════
//  UPDATE EMPLOYEE
// ════════════════════════════════════════════════════
function updateEmployee(employeeId, emp) {
  try {
    if (!employeeId) return { success: false, message: 'ไม่ระบุรหัสพนักงาน' };

    const sh      = getSheet(SHEET_EMP);
    const data    = sh.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(employeeId)) {
        Object.keys(emp).forEach(key => {
          const col = headers.indexOf(key);
          if (col !== -1) sh.getRange(i + 1, col + 1).setValue(emp[key]);
        });
        return { success: true, message: 'อัปเดตเรียบร้อยแล้ว' };
      }
    }
    return { success: false, message: 'ไม่พบพนักงาน: ' + employeeId };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ════════════════════════════════════════════════════
//  DELETE EMPLOYEE (set inactive)
// ════════════════════════════════════════════════════
function deleteEmployee(employeeId) {
  try {
    return updateEmployee(employeeId, { 'สถานะ': 'inactive' });
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ทดสอบ — รันใน Apps Script เพื่อดูค่าจริงใน Sheet
function debugRoster() {
  const sh   = getSheet(SHEET_ROSTER);
  const data = sh.getDataRange().getValues();
  const row1 = data[1]; // แถวแรกของข้อมูล
  Logger.log('row[0] value: ' + row1[0]);
  Logger.log('row[0] type: ' + typeof row1[0]);
  Logger.log('row[0] instanceof Date: ' + (row1[0] instanceof Date));
  if (row1[0] instanceof Date) {
    Logger.log('Date value: ' + row1[0].toISOString());
    const y = row1[0].getFullYear() + 543;
    const m = String(row1[0].getMonth() + 1).padStart(2,'0');
    Logger.log('Converted: ' + y + '-' + m);
  }
}
