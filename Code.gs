// ============================================================
// Code.gs — Main Server-Side Entry Point
// Complaint Handling App | Nanobank Syariah
// ============================================================

const SPREADSHEET_ID = '1iJfE0AdB0kbHAnQv2OCJRip7jNxpsgB1q8obn8dQzL4';
const SHEET_USERS      = 'users';
const SHEET_COMPLAINTS = 'complaints';
const SHEET_REPORT     = 'report';
const SHEET_MAPPING    = 'mapping';
const SHEET_PRODUCT    = 'product';
const SHEET_COPYWRITING = 'copy-writing';

// ─── Web App Entry Point ──────────────────────────────────
function doGet(e) {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Complaint Handling — Nanobank Syariah')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─── Helper ───────────────────────────────────────────────
function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh && (name === 'report' || name === 'reporting')) {
    sh = ss.getSheetByName('report') || ss.getSheetByName('reporting');
  }
  return sh;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function generateTicketId() {
  const sheet = getSheet(SHEET_COMPLAINTS);
  let maxNum = 0;

  if (sheet && sheet.getLastRow() > 1) {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let ticketColIdx = headers.findIndex(h => String(h || '').trim().toLowerCase() === 'ticket id');
    if (ticketColIdx === -1) ticketColIdx = 1; // default to Column B

    for (let i = 1; i < data.length; i++) {
      const val = String(data[i][ticketColIdx] || '').trim();
      const match = val.match(/^SC-(\d+)$/i);
      if (match) {
        const num = parseInt(match[1], 10);
        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
    }
  }

  // Juga periksa sheet report (Case ID di kolom A)
  try {
    const reportSh = getSheet(SHEET_REPORT);
    if (reportSh && reportSh.getLastRow() > 1) {
      const repData = reportSh.getDataRange().getValues();
      const repHeaders = repData[0];
      let caseColIdx = repHeaders.findIndex(h => String(h || '').trim().toLowerCase() === 'case id');
      if (caseColIdx === -1) caseColIdx = 0;

      for (let i = 1; i < repData.length; i++) {
        const val = String(repData[i][caseColIdx] || '').trim();
        const match = val.match(/^SC-(\d+)$/i);
        if (match) {
          const num = parseInt(match[1], 10);
          if (!isNaN(num) && num > maxNum) {
            maxNum = num;
          }
        }
      }
    }
  } catch (e) {
    Logger.log('Error checking report sheet: ' + e.message);
  }

  const nextNum = maxNum + 1;
  const padded = String(nextNum).padStart(5, '0');
  return `SC-${padded}`;
}

// ─── AUTH ─────────────────────────────────────────────────
function loginUser(nik, nama) {
  try {
    const sheet = getSheet(SHEET_USERS);
    const rows  = sheet.getDataRange().getValues();
    const headers = rows[0];
    const nikIdx  = headers.indexOf('nik');
    const namaIdx = headers.indexOf('nama');
    const emailIdx = headers.indexOf('email');
    const roleIdx = headers.indexOf('role');
    const picIdx  = headers.indexOf('pic');

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (String(row[nikIdx]).trim() === String(nik).trim() &&
          row[namaIdx].trim().toLowerCase() === nama.trim().toLowerCase()) {
        return {
          success: true,
          user: {
            nik:   row[nikIdx],
            nama:  row[namaIdx],
            email: row[emailIdx],
            role:  row[roleIdx],
            pic:   row[picIdx] || ''
          }
        };
      }
    }
    return { success: false, message: 'NIK atau Nama tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ─── USER MANAGEMENT ─────────────────────────────────────
function getUsers() {
  try {
    return { success: true, data: sheetToObjects(getSheet(SHEET_USERS)) };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function addUser(userData) {
  try {
    const sheet = getSheet(SHEET_USERS);
    const rows  = sheet.getDataRange().getValues();
    // Check duplicate NIK
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(userData.nik).trim()) {
        return { success: false, message: 'NIK sudah terdaftar.' };
      }
    }
    sheet.appendRow([userData.nik, userData.nama, userData.email, userData.role, userData.pic || '']);
    return { success: true, message: 'User berhasil ditambahkan.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function updateUser(userData) {
  try {
    const sheet = getSheet(SHEET_USERS);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(userData.nik).trim()) {
        sheet.getRange(i + 1, 1, 1, 5).setValues([[
          userData.nik, userData.nama, userData.email, userData.role, userData.pic || ''
        ]]);
        return { success: true, message: 'User berhasil diupdate.' };
      }
    }
    return { success: false, message: 'User tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function deleteUser(nik) {
  try {
    const sheet = getSheet(SHEET_USERS);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(nik).trim()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'User berhasil dihapus.' };
      }
    }
    return { success: false, message: 'User tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ─── PRODUCT & MAPPING DATA ───────────────────────────────
function getProductData() {
  try {
    const productRows  = sheetToObjects(getSheet(SHEET_PRODUCT));
    const mappingRows  = sheetToObjects(getSheet(SHEET_MAPPING));

    // Build unique lists
    const productCategories  = [...new Set(productRows.map(r => r['Product Category']).filter(Boolean))];
    const complaintCategories = [...new Set(productRows.map(r => r['Category']).filter(Boolean))];
    const subCategories      = [...new Set(productRows.map(r => r['Sub Category']).filter(Boolean))];
    const priorities         = [...new Set(productRows.map(r => r['Priority']).filter(Boolean))];

    return {
      success: true,
      productCategories,
      complaintCategories,
      subCategories,
      priorities,
      mappingData: mappingRows
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ─── COMPLAINT SUBMISSION ─────────────────────────────────
function submitComplaint(formData) {
  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(30000);
  if (!hasLock) {
    return { success: false, message: 'Server sedang sibuk memproses antrean complaint lain. Silakan coba sesaat lagi.' };
  }

  try {
    const sheet     = getSheet(SHEET_COMPLAINTS);
    const reportSh  = getSheet(SHEET_REPORT);
    const ticketId  = generateTicketId();
    const timestamp = new Date();

    // Build complaints row (matching sheet headers)
    const complaintRow = [
      timestamp,
      ticketId,
      formData.productCategory,
      formData.complaintCategory,
      formData.subCategory,
      formData.sla,
      formData.codeId,
      formData.priority,
      formData.classification,
      formData.reason,
      formData.name,
      formData.accountNumber,
      formData.atmCardNumber,
      formData.transactionAmount1,
      formData.transactionAmount2,
      formData.emailAddress,
      formData.phoneNumber,
      formData.dateTransaction,
      formData.destinationPhoneNumber,
      formData.providerName,
      formData.pic,
      formData.emailPic,
      formData.description
    ];
    sheet.appendRow(complaintRow);

    // Build report row (matching sheet 'report' headers dynamically)
    const dueDate = new Date(timestamp);
    dueDate.setDate(dueDate.getDate() + Number(formData.sla || 0));

    const reportHeaders = reportSh.getRange(1, 1, 1, Math.max(reportSh.getLastColumn(), 24)).getValues()[0];
    const reportValueMap = {
      'case id': ticketId,
      'create date/time': timestamp,
      'sla': formData.sla,
      'monitor by': '',
      'account name': formData.name,
      'code id sub-category': formData.codeId,
      'sub category': formData.subCategory,
      'created by': formData.createdBy,
      'product category': formData.productCategory,
      'category': formData.complaintCategory,
      'complaint handling status': 'Open',
      'currentpic': formData.pic,
      'emailpic': formData.emailPic,
      'cif': '',
      'account number': formData.accountNumber,
      'classification': formData.classification,
      'reason': formData.reason,
      'date completed': '',
      'due date': dueDate,
      'nominal': formData.transactionAmount1,
      'tanggal transaksi': formData.dateTransaction,
      'description report': formData.description,
      'location and merchant name': '',
      'completeday': ''
    };

    const reportRow = reportHeaders.map(h => {
      const key = String(h || '').trim().toLowerCase();
      return reportValueMap.hasOwnProperty(key) ? reportValueMap[key] : '';
    });
    reportSh.appendRow(reportRow);

    // Send emails
    _sendEmails(formData, ticketId);

    return { success: true, ticketId };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function _sendEmails(formData, ticketId) {
  const cwRows = sheetToObjects(getSheet(SHEET_COPYWRITING));
  const template = cwRows.find(r => String(r['No']) === '1');
  if (!template) return;

  function fillTemplate(text) {
    return text
      .replace(/\{Ticket ID\}/g, ticketId)
      .replace(/\{SLA\}/g, formData.sla || '-');
  }

  // Parse subject + body from copywriting cell
  function parseEmailContent(raw) {
    const lines  = raw.split('\n');
    const subjectLine = lines.find(l => l.startsWith('Subject:'));
    const subject = subjectLine ? subjectLine.replace('Subject:', '').trim() : 'Notifikasi Complaint';
    const bodyStart = raw.indexOf('Isi:');
    const body = bodyStart !== -1 ? raw.slice(bodyStart + 4).trim() : raw;
    return { subject: fillTemplate(subject), body: fillTemplate(body) };
  }

  // Email to customer
  if (formData.emailAddress) {
    try {
      const { subject, body } = parseEmailContent(template['Copywrite Kirim ke Nasabah']);
      MailApp.sendEmail({ to: formData.emailAddress, subject, body });
    } catch (e) { Logger.log('Email nasabah gagal: ' + e.message); }
  }

  // Email to PK
  if (formData.emailPic) {
    try {
      const { subject, body } = parseEmailContent(template['Copywrite Kirim ke PK']);
      MailApp.sendEmail({ to: formData.emailPic, subject, body });
    } catch (e) { Logger.log('Email PK gagal: ' + e.message); }
  }
}

// ─── REPORTING ────────────────────────────────────────────
function getReportData() {
  try {
    const sheet = getSheet(SHEET_REPORT);
    if (!sheet) {
      return { success: false, message: 'Sheet report / reporting tidak ditemukan.' };
    }
    const values = sheet.getDataRange().getValues();
    if (!values || values.length === 0) {
      return { success: true, headers: [], data: [] };
    }

    const rawHeaders = values[0].map(h => String(h != null ? h : ''));
    const cleanHeaders = rawHeaders.map(h => h.trim());

    const rows = values.length > 1 ? values.slice(1).map(row => {
      const obj = {};
      rawHeaders.forEach((rawH, i) => {
        let val = row[i];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, 'Asia/Jakarta', "yyyy-MM-dd'T'HH:mm:ss");
        }
        obj[rawH] = val != null ? val : '';
        obj[cleanHeaders[i]] = val != null ? val : '';
      });
      return obj;
    }) : [];

    return {
      success: true,
      headers: cleanHeaders.filter(h => h.length > 0),
      rawHeaders: rawHeaders,
      data: rows
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}
