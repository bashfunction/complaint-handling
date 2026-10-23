// ============================================================
// Code.gs — Main Server-Side Entry Point
// Complaint Handling App | Nanobank Syariah
// ============================================================

const SPREADSHEET_ID    = '1iJfE0AdB0kbHAnQv2OCJRip7jNxpsgB1q8obn8dQzL4';
const SHEET_USERS       = 'users';
const SHEET_COMPLAINTS  = 'complaints';
const SHEET_REPORT      = 'report';
const SHEET_MAPPING     = 'mapping';
const SHEET_PRODUCT     = 'product';
const SHEET_COPYWRITING = 'copy-writing';
const GFORM_ID          = '1qrczQsuAq-VA-nF9L9-bIwrQ_zrMfirchT-4Q0gO5oo';

// ─── Web App Entry Point ──────────────────────────────────
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.initialAuth = getCurrentUserProfile();
  return template
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

// ─── AUTH / CHROME SSO ────────────────────────────────────
function getCurrentUserProfile() {
  try {
    let email = '';
    try {
      email = Session.getActiveUser().getEmail();
    } catch (e) {}
    if (!email) {
      try {
        email = Session.getEffectiveUser().getEmail();
      } catch (e) {}
    }

    const sheet = getSheet(SHEET_USERS);
    if (!sheet) {
      return { success: false, message: 'Sheet users tidak ditemukan.' };
    }

    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) {
      return { success: false, message: 'Data users pada sheet kosong.' };
    }

    const headers = rows[0].map(h => String(h || '').trim().toLowerCase());
    const emailIdx = headers.indexOf('email');
    const namaIdx = headers.indexOf('nama');
    const nikIdx = headers.indexOf('nik');
    const roleIdx = headers.indexOf('role');
    const picIdx = headers.indexOf('pic');

    const cleanEmail = String(email || '').trim().toLowerCase();

    // 1. Search user by Chrome profile email
    if (cleanEmail) {
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const rowEmail = String(row[emailIdx] || '').trim().toLowerCase();
        if (rowEmail === cleanEmail) {
          return {
            success: true,
            user: {
              nik: String(row[nikIdx] || ''),
              nama: String(row[namaIdx] || ''),
              email: String(row[emailIdx] || email),
              role: String(row[roleIdx] || 'cs').toLowerCase(),
              pic: String(row[picIdx] || '')
            }
          };
        }
      }
    }

    // 2. Email detected from Chrome, but not registered in users sheet
    if (cleanEmail) {
      return {
        success: false,
        email: email,
        message: `Akun Google (${email}) belum terdaftar pada sistem Complaint Handling. Hubungi Administrator untuk mendaftarkan akun Anda.`
      };
    }

    // 3. Fallback: if Session.getActiveUser() is empty (e.g. running in testing without signed-in user or iframe restrictions)
    // Fallback to first user in sheet so the app still functions seamlessly during testing:
    if (rows.length > 1) {
      const row = rows[1];
      return {
        success: true,
        isFallback: true,
        user: {
          nik: String(row[nikIdx] || ''),
          nama: String(row[namaIdx] || ''),
          email: String(row[emailIdx] || ''),
          role: String(row[roleIdx] || 'admin').toLowerCase(),
          pic: String(row[picIdx] || '')
        }
      };
    }

    return {
      success: false,
      message: 'Tidak dapat mendeteksi akun Google Chrome. Pastikan Anda telah login ke akun Google di browser Chrome.'
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function loginUser(nik, nama) {
  return getCurrentUserProfile();
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

    const userRows = sheetToObjects(getSheet(SHEET_USERS));
    const picUsers = userRows.map(u => ({
      nama: u.nama || '',
      email: u.email || '',
      pic: u.pic || ''
    }));

    return {
      success: true,
      productCategories,
      complaintCategories,
      subCategories,
      priorities,
      mappingData: mappingRows,
      picUsers
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

    // Build report row — Due Date dihitung dalam hari kerja (exclude Sabtu & Minggu)
    const dueDate = addWorkingDays_(timestamp, Number(formData.sla || 0));

    const reportHeaders = reportSh.getRange(1, 1, 1, Math.max(reportSh.getLastColumn(), 24)).getValues()[0];
    const reportValueMap = {
      'case id':                    ticketId,
      'create date/time':           timestamp,
      'sla':                        formData.sla,
      'monitor by':                 '',
      'account name':               formData.name,
      'code id sub-category':       formData.codeId,
      'sub category':               formData.subCategory,
      'created by':                 formData.createdBy,
      'product category':           formData.productCategory,
      'category':                   formData.complaintCategory,
      'complaint handling status':  'Open',
      'currentpic':                 formData.pic,
      'emailpic':                   formData.emailPic,
      'cif':                        '',
      'account number':             formData.accountNumber,
      'classification':             formData.classification,
      'reason':                     formData.reason,
      'date completed':             '',
      'due date':                   dueDate,
      'nominal':                    formData.transactionAmount1,
      'tanggal transaksi':          formData.dateTransaction,
      'description report':         formData.description,
      'location and merchant name': '',
      'completeday':                ''
    };

    const reportRow = reportHeaders.map(h => {
      const key = String(h || '').trim().toLowerCase();
      return reportValueMap.hasOwnProperty(key) ? reportValueMap[key] : '';
    });
    reportSh.appendRow(reportRow);

    // Send emails on create
    _sendEmailsOnCreate_(formData, ticketId, dueDate);

    return { success: true, ticketId };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ─── EMAIL: SAAT TIKET DIBUAT ─────────────────────────────
function _sendEmailsOnCreate_(formData, ticketId, dueDate) {
  const cwRows   = sheetToObjects(getSheet(SHEET_COPYWRITING));
  const template = cwRows.find(r => String(r['No'] || '').trim() === '1');
  if (!template) return;

  const tz        = 'Asia/Jakarta';
  const tanggalStr = Utilities.formatDate(new Date(), tz, 'dd MMMM yyyy HH:mm') + ' WIB';
  const dueDateStr = Utilities.formatDate(dueDate, tz, 'dd MMMM yyyy');

  function fill(text) {
    if (!text) return '';
    return text
      .replace(/\{Ticket ID\}/g, ticketId)
      .replace(/\{SLA\}/g,         formData.sla || '-')
      .replace(/\{Tanggal\}/g,     tanggalStr)
      .replace(/\{Product\}/g,     formData.productCategory || '-')
      .replace(/\{Sub Category\}/g, formData.subCategory || '-')
      .replace(/\{PIC\}/g,         formData.pic || '-')
      .replace(/\{Deskripsi\}/g,   formData.description || '-')
      .replace(/\{Due Date\}/g,    dueDateStr);
  }

  function parseEmail(raw) {
    if (!raw) return null;
    const subjectLine = raw.split('\n').find(l => l.startsWith('Subject:'));
    const subject = subjectLine ? subjectLine.replace('Subject:', '').trim() : 'Notifikasi Complaint';
    const bodyStart = raw.indexOf('Isi:');
    const body = bodyStart !== -1 ? raw.slice(bodyStart + 4).trim() : raw;
    return { subject: fill(subject), body: fill(body) };
  }

  // 1. Email ke Nasabah
  if (formData.emailAddress) {
    try {
      const e = parseEmail(template['Copywrite Kirim ke Nasabah']);
      if (e) MailApp.sendEmail({ to: formData.emailAddress, subject: e.subject, body: e.body, name: 'Nanobank Syariah' });
    } catch (err) { Logger.log('Email nasabah gagal: ' + err.message); }
  }

  // 2. Email ke PK
  if (formData.emailPic) {
    try {
      const e = parseEmail(template['Copywrite Kirim ke PK']);
      if (e) MailApp.sendEmail({ to: formData.emailPic, subject: e.subject, body: e.body, name: 'Nanobank Syariah' });
    } catch (err) { Logger.log('Email PK gagal: ' + err.message); }
  }

  // 3. Email ke UK dengan tombol 'Konfirmasi' (prefilled Task ID & Completed)
  if (formData.emailPic) {
    try {
      const cwUK = template['Copywrite kirim ke UK'];
      const e    = parseEmail(cwUK);
      if (e) {
        const confirmUrl     = _buildPrefillUrl_(ticketId, 'Completed');
        const confirmBtnHtml = _buildConfirmButton_(ticketId);

        // Plain text version
        const bodyText = e.body
          .replace(/\[Tombol\s*['"]?Konfirmasi['"]?\]/gi, 'Link Konfirmasi: ' + confirmUrl)
          .replace(/\[Link GForm Anda\]/g, confirmUrl)
          .replace(/Link Konfirmasi Status:.*$/m, 'Link Konfirmasi Status: ' + confirmUrl);

        // HTML version
        let htmlBody = e.body.replace(/\n/g, '<br>');
        if (/\[Tombol\s*['"]?Konfirmasi['"]?\]/i.test(htmlBody)) {
          htmlBody = htmlBody.replace(/\[Tombol\s*['"]?Konfirmasi['"]?\]/gi, confirmBtnHtml);
        } else {
          htmlBody = htmlBody
            .replace(/\[Link GForm Anda\]/g, confirmUrl)
            .replace(/Link Konfirmasi Status:.*$/m, 'Link Konfirmasi Status: ' + confirmUrl);
          htmlBody += '<br><br>' + confirmBtnHtml;
        }

        MailApp.sendEmail({
          to:       formData.emailPic,
          subject:  e.subject,
          body:     bodyText,
          htmlBody: htmlBody,
          name:     'Customer Care Specialist - Nanobank Syariah'
        });
      }
    } catch (err) { Logger.log('Email UK gagal: ' + err.message); }
  }
}

// ─── GFORM PREFILL LINK BUILDER ───────────────────────────
function _buildPrefillUrl_(taskId, statusValue) {
  try {
    const form  = FormApp.openById(GFORM_ID);
    const items = form.getItems();
    const findItem = title => items.find(x => x.getTitle().trim() === title || x.getTitle().trim() === title.trim());

    const taskIdItem = findItem('Task ID');
    const statusItem = findItem('Status');
    const resp = form.createResponse();

    if (taskIdItem) {
      resp.withItemResponse(taskIdItem.asTextItem().createResponse(taskId));
    }
    if (statusValue && statusItem) {
      try {
        const sResp = statusItem.getType() === FormApp.ItemType.LIST
          ? statusItem.asListItem().createResponse(statusValue)
          : statusItem.asMultipleChoiceItem().createResponse(statusValue);
        resp.withItemResponse(sResp);
      } catch (e) { Logger.log('Status prefill error: ' + e.message); }
    }
    return resp.toPrefilledUrl();
  } catch (err) {
    Logger.log('buildPrefillUrl error: ' + err.message);
    return 'https://docs.google.com/forms/d/' + GFORM_ID + '/viewform';
  }
}

// ─── HTML ACTION BUTTON FOR UK EMAIL ─────────────────────
function _buildConfirmButton_(ticketId) {
  const urlConfirm = _buildPrefillUrl_(ticketId, 'Completed');
  return '<div style="margin:20px 0;">' +
    `<a href="${urlConfirm}" target="_blank" style="display:inline-block;padding:12px 32px;` +
    `background:#002657;color:#ffffff;text-decoration:none;border-radius:8px;font-weight:bold;` +
    `font-size:14px;font-family:Arial,sans-serif;box-shadow:0 2px 6px rgba(0,0,0,0.2);border:1px solid #001a3d;">` +
    `&#9998; Konfirmasi` +
    `</a>` +
    '</div>';
}

function _buildActionButtons_(ticketId) {
  return _buildConfirmButton_(ticketId);
}

// ─── TRIGGER: UK SUBMIT GFORM ─────────────────────────────
// Pasang trigger: Triggers > Add Trigger > onGFormSubmit
// From: Spreadsheet, Event: On form submit
function onGFormSubmit(e) {
  try {
    const nv        = e.namedValues;
    const taskId    = String((nv['Task ID'] || nv['Task ID\n'] || [''])[0]).trim();
    const status    = String((nv['Status']  || [''])[0]).trim();
    const alasan    = String((nv['Alasan Extend'] || [''])[0]).trim();
    const slaExt    = String((nv['SLA']     || [''])[0]).trim();
    const submitter = String((nv['Email Address'] || [''])[0]).trim();
    const timestamp = new Date();

    if (!taskId || !status) return;

    const reportSh = getSheet(SHEET_REPORT);
    if (!reportSh) return;

    const data    = reportSh.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());

    const colCaseId   = headers.indexOf('case id');
    const colStatus   = headers.indexOf('complaint handling status');
    const colDateComp = headers.indexOf('date completed');
    const colDueDate  = headers.indexOf('due date');
    const colComplete = headers.indexOf('completeday');
    const colCreateDt = headers.indexOf('create date/time');
    const colPIC      = headers.indexOf('currentpic');
    const colEmailPic = headers.indexOf('emailpic');
    const colProdCat  = headers.indexOf('product category');
    const colSubCat   = headers.indexOf('sub category');
    const colSLA      = headers.indexOf('sla');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colCaseId] || '').trim() !== taskId) continue;
      const row = data[i];

      if (status === 'Completed') {
        const createDt    = row[colCreateDt] instanceof Date ? row[colCreateDt] : new Date(row[colCreateDt]);
        const completeDay = _countWorkingDays_(createDt, timestamp);
        if (colStatus   >= 0) reportSh.getRange(i + 1, colStatus   + 1).setValue('Completed');
        if (colDateComp >= 0) reportSh.getRange(i + 1, colDateComp + 1).setValue(timestamp);
        if (colComplete >= 0) reportSh.getRange(i + 1, colComplete + 1).setValue(completeDay);
        _sendToCCS_(taskId, status, submitter, alasan, slaExt, row, headers, timestamp);

      } else if (status === 'Extend') {
        const currentDue  = row[colDueDate] instanceof Date ? row[colDueDate] : new Date(row[colDueDate] || new Date());
        const slaExtNum   = parseInt(slaExt || '10', 10) || 10;
        const newDueDate  = addWorkingDays_(currentDue, slaExtNum);
        if (colStatus  >= 0) reportSh.getRange(i + 1, colStatus  + 1).setValue('Extend');
        if (colDueDate >= 0) reportSh.getRange(i + 1, colDueDate + 1).setValue(newDueDate);
        _sendToCCS_(taskId, status, submitter, alasan, slaExt, row, headers, timestamp);
        _sendExtendNotifToNasabah_(taskId, newDueDate);
        _sendExtendNotifToUK_(taskId, alasan, slaExtNum, row, headers);
        _sendExtendNotifToPK_(taskId, row, headers);
      }
      break;
    }
  } catch (err) {
    Logger.log('onGFormSubmit error: ' + err.message);
  }
}

// ─── NOTIFIKASI CCS ──────────────────────────────────────
function _sendToCCS_(taskId, status, submitter, alasan, slaExt, row, headers, timestamp) {
  try {
    const cwRows   = sheetToObjects(getSheet(SHEET_COPYWRITING));
    const template = cwRows.find(r => String(r['No'] || '').trim() === '2');
    if (!template) return;

    const tz     = 'Asia/Jakarta';
    const colPIC = headers.indexOf('currentpic');
    const tsStr  = Utilities.formatDate(timestamp, tz, 'dd MMMM yyyy HH:mm') + ' WIB';
    const raw    = String(template['Copywrite kirim ke CCS'] || '');
    const filled = raw
      .replace(/\{Ticket ID\}/g,     taskId)
      .replace(/\{PIC\}/g,           colPIC >= 0 ? String(row[colPIC] || '-') : '-')
      .replace(/\{Timestamp\}/g,     tsStr)
      .replace(/\{Email Pengisi\}/g, submitter || '-')
      .replace(/\{Status\}/g,        status)
      .replace(/\{Alasan Extend\}/g, alasan || '-')
      .replace(/\{SLA Extend\}/g,    slaExt || '-');

    const subjectLine = filled.split('\n').find(l => l.startsWith('Subject:'));
    const subject     = subjectLine ? subjectLine.replace('Subject:', '').trim() : '[CCS] Update Tiket ' + taskId;
    const bodyStart   = filled.indexOf('Isi:');
    const body        = bodyStart !== -1 ? filled.slice(bodyStart + 4).trim() : filled;

    const users     = sheetToObjects(getSheet(SHEET_USERS));
    const ccsEmails = users
      .filter(u => String(u.role || '').trim().toLowerCase() === 'ccs' && u.email)
      .map(u => u.email);
    if (ccsEmails.length > 0) {
      MailApp.sendEmail({ to: ccsEmails.join(','), subject, body, name: 'Sistem Complaint Handling' });
    }
  } catch (err) { Logger.log('sendToCCS error: ' + err.message); }
}

// ─── NOTIFIKASI EXTEND KE NASABAH ────────────────────────
function _sendExtendNotifToNasabah_(taskId, newDueDate) {
  try {
    const cwRows   = sheetToObjects(getSheet(SHEET_COPYWRITING));
    const template = cwRows.find(r => String(r['No'] || '').trim() === '2');
    if (!template) return;

    const compSh      = getSheet(SHEET_COMPLAINTS);
    const compData    = compSh.getDataRange().getValues();
    const compHeaders = compData[0].map(h => String(h || '').trim().toLowerCase());
    const ticketIdx   = compHeaders.findIndex(h => h === 'ticket id');
    const emailIdx    = compHeaders.findIndex(h => h === 'email address');
    let nasabahEmail  = '';

    for (let i = 1; i < compData.length; i++) {
      if (String(compData[i][ticketIdx] || '').trim() === taskId) {
        nasabahEmail = String(compData[i][emailIdx] || '');
        break;
      }
    }
    if (!nasabahEmail) return;

    const tz         = 'Asia/Jakarta';
    const dueDateStr = Utilities.formatDate(newDueDate, tz, 'dd MMMM yyyy');
    const raw        = String(template['Copywrite Kirim ke Nasabah'] || '');
    const filled     = raw.replace(/\{Ticket ID\}/g, taskId).replace(/\{Due Date\}/g, dueDateStr);

    const subjectLine = filled.split('\n').find(l => l.startsWith('Subject:'));
    const subject     = subjectLine ? subjectLine.replace('Subject:', '').trim() : 'Perpanjangan Tiket ' + taskId;
    const bodyStart   = filled.indexOf('Isi:');
    const body        = bodyStart !== -1 ? filled.slice(bodyStart + 4).trim() : filled;
    MailApp.sendEmail({ to: nasabahEmail, subject, body, name: 'Nanobank Syariah' });
  } catch (err) { Logger.log('sendExtendNotifToNasabah error: ' + err.message); }
}

// ─── NOTIFIKASI EXTEND KE UK ─────────────────────────────
function _sendExtendNotifToUK_(taskId, alasan, slaExtNum, row, headers) {
  try {
    const cwRows   = sheetToObjects(getSheet(SHEET_COPYWRITING));
    const template = cwRows.find(r => String(r['No'] || '').trim() === '2');
    if (!template) return;

    const colEmailPic = headers.indexOf('emailpic');
    const colPIC      = headers.indexOf('currentpic');
    const colProdCat  = headers.indexOf('product category');
    const colSubCat   = headers.indexOf('sub category');
    const emailUK     = colEmailPic >= 0 ? String(row[colEmailPic] || '') : '';
    if (!emailUK) return;

    const gformLink = _buildPrefillUrl_(taskId, '');
    const raw = String(template['Copywrite kirim ke UK'] || '');
    const filled = raw
      .replace(/\{Ticket ID\}/g,     taskId)
      .replace(/\{PIC\}/g,           colPIC     >= 0 ? String(row[colPIC]     || '-') : '-')
      .replace(/\{Product\}/g,       colProdCat >= 0 ? String(row[colProdCat] || '-') : '-')
      .replace(/\{Sub Category\}/g,  colSubCat  >= 0 ? String(row[colSubCat]  || '-') : '-')
      .replace(/\{Alasan Extend\}/g, alasan || '-')
      .replace(/\[Link GForm Anda\]/g, gformLink)
      .replace(/Link Konfirmasi Status:.*$/m, 'Link Konfirmasi Status: ' + gformLink);

    const subjectLine = filled.split('\n').find(l => l.startsWith('Subject:'));
    const subject     = subjectLine ? subjectLine.replace('Subject:', '').trim() : 'Extend SLA - Tiket ' + taskId;
    const bodyStart   = filled.indexOf('Isi:');
    const body        = bodyStart !== -1 ? filled.slice(bodyStart + 4).trim() : filled;
    const htmlBody    = body.replace(/\n/g, '<br>') + '<br><br>' + _buildActionButtons_(taskId);
    MailApp.sendEmail({ to: emailUK, subject, body, htmlBody, name: 'Customer Care Specialist - Nanobank Syariah' });
  } catch (err) { Logger.log('sendExtendNotifToUK error: ' + err.message); }
}

// ─── NOTIFIKASI EXTEND KE PK ─────────────────────────────
function _sendExtendNotifToPK_(taskId, row, headers) {
  try {
    const cwRows   = sheetToObjects(getSheet(SHEET_COPYWRITING));
    const template = cwRows.find(r => String(r['No'] || '').trim() === '2');
    if (!template) return;

    const colEmailPic = headers.indexOf('emailpic');
    const emailPK     = colEmailPic >= 0 ? String(row[colEmailPic] || '') : '';
    if (!emailPK) return;

    const raw    = String(template['Copywrite Kirim ke PK'] || '');
    const filled = raw.replace(/\{Ticket ID\}/g, taskId);
    const subjectLine = filled.split('\n').find(l => l.startsWith('Subject:'));
    const subject     = subjectLine ? subjectLine.replace('Subject:', '').trim() : 'Perpanjangan SLA - Tiket ' + taskId;
    const bodyStart   = filled.indexOf('Isi:');
    const body        = bodyStart !== -1 ? filled.slice(bodyStart + 4).trim() : filled;
    MailApp.sendEmail({ to: emailPK, subject, body, name: 'Nanobank Syariah' });
  } catch (err) { Logger.log('sendExtendNotifToPK error: ' + err.message); }
}

// ─── TRIGGER HARIAN: REMINDER H-0 SLA ────────────────────
// Kirim email ke UK jika tiket Open/Extend dan Due Date = hari ini
// Daftarkan via addTriggers() sekali dari GAS Editor
function sendSLAReminderEmails() {
  try {
    const reportSh = getSheet(SHEET_REPORT);
    if (!reportSh) return;
    const data    = reportSh.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());

    const colStatus   = headers.indexOf('complaint handling status');
    const colDueDate  = headers.indexOf('due date');
    const colCaseId   = headers.indexOf('case id');
    const colEmailPic = headers.indexOf('emailpic');
    const colPIC      = headers.indexOf('currentpic');
    const colProdCat  = headers.indexOf('product category');
    const colSubCat   = headers.indexOf('sub category');
    const colSLA      = headers.indexOf('sla');
    const colCreateDt = headers.indexOf('create date/time');

    const cwRows   = sheetToObjects(getSheet(SHEET_COPYWRITING));
    const template = cwRows.find(r => String(r['No'] || '').trim() === '3');
    if (!template) return;

    const tz       = 'Asia/Jakarta';
    const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    for (let i = 1; i < data.length; i++) {
      const row    = data[i];
      const status = String(row[colStatus] || '').trim();
      if (status !== 'Open' && status !== 'Extend') continue;

      const dueDate    = row[colDueDate] instanceof Date ? row[colDueDate] : new Date(row[colDueDate]);
      const dueDateStr = Utilities.formatDate(dueDate, tz, 'yyyy-MM-dd');
      if (dueDateStr !== todayStr) continue;

      const taskId  = String(row[colCaseId]   || '');
      const emailUK = String(row[colEmailPic] || '');
      if (!taskId || !emailUK) continue;

      const createDt   = row[colCreateDt] instanceof Date ? row[colCreateDt] : new Date(row[colCreateDt]);
      const gformLink  = _buildPrefillUrl_(taskId, '');
      const raw        = String(template['Copywrite kirim ke UK'] || '');
      const filled     = raw
        .replace(/\{Ticket ID\}/g,    taskId)
        .replace(/\{Tanggal\}/g,      Utilities.formatDate(createDt, tz, 'dd MMMM yyyy'))
        .replace(/\{PIC\}/g,          colPIC     >= 0 ? String(row[colPIC]     || '-') : '-')
        .replace(/\{Product\}/g,      colProdCat >= 0 ? String(row[colProdCat] || '-') : '-')
        .replace(/\{Sub Category\}/g, colSubCat  >= 0 ? String(row[colSubCat]  || '-') : '-')
        .replace(/\{SLA\}/g,          colSLA     >= 0 ? String(row[colSLA]     || '-') : '-')
        .replace(/\{Due Date\}/g,     Utilities.formatDate(dueDate, tz, 'dd MMMM yyyy'))
        .replace(/\[Link GForm Anda\]/g, gformLink)
        .replace(/Link Konfirmasi Status:.*$/m, 'Link Konfirmasi Status: ' + gformLink);

      const subjectLine = filled.split('\n').find(l => l.startsWith('Subject:'));
      const subject     = subjectLine ? subjectLine.replace('Subject:', '').trim() : '[REMINDER H-0] Tiket ' + taskId;
      const bodyStart   = filled.indexOf('Isi:');
      const body        = bodyStart !== -1 ? filled.slice(bodyStart + 4).trim() : filled;
      const htmlBody    = body.replace(/\n/g, '<br>') + '<br><br>' + _buildActionButtons_(taskId);
      try {
        MailApp.sendEmail({ to: emailUK, subject, body, htmlBody, name: 'Customer Care Specialist - Nanobank Syariah' });
      } catch (err) {
        Logger.log('Reminder gagal untuk ' + taskId + ': ' + err.message);
      }
    }
  } catch (err) {
    Logger.log('sendSLAReminderEmails error: ' + err.message);
  }
}

// ─── SETUP TRIGGER (Jalankan SEKALI dari GAS Editor) ─────
function addTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendSLAReminderEmails') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendSLAReminderEmails')
    .timeBased().everyDays(1).atHour(7).create();
  Logger.log('Trigger sendSLAReminderEmails berhasil didaftarkan: setiap hari pukul 07.00');
}

// ─── HELPER: HITUNG HARI KERJA ────────────────────────────
// Tambahkan N hari kerja (exclude Sabtu=6 & Minggu=0)
function addWorkingDays_(startDate, days) {
  const result = new Date(startDate);
  let added = 0;
  while (added < days) {
    result.setDate(result.getDate() + 1);
    const dow = result.getDay();
    if (dow !== 0 && dow !== 6) added++;
  }
  return result;
}

// Hitung hari kerja antara dua tanggal
function _countWorkingDays_(startDate, endDate) {
  let count = 0;
  const cur = new Date(startDate);
  cur.setHours(0, 0, 0, 0);
  const end = new Date(endDate);
  end.setHours(0, 0, 0, 0);
  while (cur < end) {
    cur.setDate(cur.getDate() + 1);
    const dow = cur.getDay();
    if (dow !== 0 && dow !== 6) count++;
  }
  return count;
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
