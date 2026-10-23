/**
 * TOMBOL "COMPLETE" & "RE-ASSIGN" -> Google Form prefilled -> Google Sheet
 *
 * Sheet "Tasks"  : A = Task ID, B = Nama Task, C = Assignee (email), D = Deadline, E = Status
 * Form fields    : "Task ID"   (Short answer)
 *                  "Status"    (Multiple choice / Dropdown: Complete, Re-assign)
 * Form dihubungkan ke spreadsheet yang sama (Responses > Link to Sheets).
 */

const CONFIG = {
  FORM_ID: '1Jyg4QnhyjNXnqqrydHEdRUAR8zNJsONyxOTLI4_OOFM',
  TASKS_SHEET: 'Tasks',
  Q_TASK_ID: 'Task ID',
  Q_STATUS: 'Status',
  OPT_COMPLETE: 'Complete',
  OPT_REASSIGN: 'Re-assign',
  SENDER_NAME: 'Task Manager'
};

// ====== KIRIM TASK (hanya yang statusnya kosong / Pending) ======
function sendTasks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const rows = sheet.getDataRange().getValues();

  const urlComplete = buildUrlTemplate_(CONFIG.OPT_COMPLETE);
  const urlReassign = buildUrlTemplate_(CONFIG.OPT_REASSIGN);

  for (let i = 1; i < rows.length; i++) {
    const [taskId, taskName, assignee, deadline, status] = rows[i];
    if (!taskId || !assignee) continue;
    if (status && String(status).toLowerCase() !== 'pending') continue;

    sendTaskEmail_(taskId, taskName, assignee, deadline, urlComplete, urlReassign);
    sheet.getRange(i + 1, 5).setValue('Sent');
  }
}

function sendTaskEmail_(taskId, taskName, assignee, deadline, urlComplete, urlReassign) {
  const link = tpl => tpl.replace('__ID__', encodeURIComponent(taskId));
  const dl = deadline instanceof Date
    ? Utilities.formatDate(deadline, Session.getScriptTimeZone(), 'dd MMM yyyy')
    : (deadline || '-');

  const btn = (url, label, color) =>
    `<a href="${url}" target="_blank" style="display:inline-block;padding:12px 24px;margin-right:8px;
      background:${color};color:#fff;text-decoration:none;border-radius:6px;font-weight:bold;
      font-size:14px;">${label}</a>`;

  const html = `
  <div style="font-family:Arial,sans-serif;max-width:600px;margin:auto;color:#222;">
    <h2>Task: ${taskName}</h2>
    <p><b>Task ID:</b> ${taskId}<br><b>Deadline:</b> ${dl}</p>
    <p>Mohon konfirmasi status task ini:</p>
    <p>
      ${btn(link(urlComplete), 'Complete', '#1e8e3e')}
      ${btn(link(urlReassign), 'Re-assign', '#e37400')}
    </p>
  </div>`;

  MailApp.sendEmail({
    to: assignee,
    subject: '[' + taskId + '] ' + taskName,
    name: CONFIG.SENDER_NAME,
    body: 'Task ' + taskId + ' - ' + taskName + '. Buka email ini dalam format HTML untuk melihat tombol.',
    htmlBody: html
  });
}

// ====== BUAT TEMPLATE LINK PREFILL (Task ID = placeholder, Status = tetap) ======
function buildUrlTemplate_(statusValue) {
  const form = FormApp.openById(CONFIG.FORM_ID);
  const items = form.getItems();
  const find = title => {
    const it = items.find(x => x.getTitle() === title);
    if (!it) throw new Error('Pertanyaan "' + title + '" tidak ditemukan di Form.');
    return it;
  };

  const resp = form.createResponse();
  resp.withItemResponse(find(CONFIG.Q_TASK_ID).asTextItem().createResponse('__ID__'));

  const statusItem = find(CONFIG.Q_STATUS);
  const statusResp = statusItem.getType() === FormApp.ItemType.LIST
    ? statusItem.asListItem().createResponse(statusValue)
    : statusItem.asMultipleChoiceItem().createResponse(statusValue);
  resp.withItemResponse(statusResp);

  return resp.toPrefilledUrl();
}

// ====== TRIGGER (OPSIONAL): SAAT FORM DISUBMIT -> UPDATE STATUS DI SHEET TASKS ======
// Data jawaban sudah otomatis tersimpan di tab respons Form.
// Trigger ini hanya menyinkronkan kolom Status di tab "Tasks".
// Triggers > Add Trigger > onFormSubmit > From spreadsheet > On form submit
function onFormSubmit(e) {
  const nv = e.namedValues;
  const taskId = String((nv[CONFIG.Q_TASK_ID] || [''])[0]).trim();
  const status = String((nv[CONFIG.Q_STATUS] || [''])[0]).trim();
  if (!taskId || !status) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TASKS_SHEET);
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === taskId) {
      sheet.getRange(i + 1, 5).setValue(
        status === CONFIG.OPT_COMPLETE ? 'Completed' : 'Re-assign'
      );
      break;
    }
  }
}