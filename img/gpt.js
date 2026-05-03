// =============== CONFIG =================
var MASTER_FOLDER_ID = '1ojvXheEI1odL6WrLTX2g-1wXLSSR16f-';
var LOGO_ID = '1j3FE5Ym_w2LyUvmRj0nlLkQixsGPUUkI';
var QR_ID   = '109aBIRn-fzkWgyELMQA4j3HCvakrU-yP';
var RECEIPT_EMAIL = 'bijoyhaque8@gmail.com';
// =======================================

function onFormSubmit(e) {

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formSheet = ss.getSheetByName('Form responses_AdmissionForm');
    var registerSheet = ss.getSheetByName('ID & Register Sheet');

    var row = e.range.getRow();
    var data = formSheet.getRange(row, 1, 1, 15).getValues()[0];

    // ===== FORM DATA =====
    var timestamp   = data[0];
    var name        = data[1];
    var phone       = String(data[2]).trim();
    var course      = data[3];
    var batch       = data[4];
    var courseFee   = Number(data[6]) || 0;
    var paidNow     = Number(data[7]) || 0;

    var receiptDate = Utilities.formatDate(
      new Date(timestamp),
      Session.getScriptTimeZone(),
      "dd MMM yyyy"
    );

    // ===== TOTAL PAID (FROM RESPONSE SHEET) =====
    var lastRow = formSheet.getLastRow();
    var allData = formSheet.getRange(2, 1, lastRow - 1, 8).getValues();

    var totalPaid = 0;
    var originalCourseFee = courseFee;

    allData.forEach(function(r) {
      if (String(r[2]).trim() === phone) {
        totalPaid += Number(r[7]) || 0;
        if (Number(r[6])) {
          originalCourseFee = Number(r[6]);
        }
      }
    });

    var dueAmount = originalCourseFee - totalPaid;
    var status = dueAmount <= 0 ? "Full Paid" : "Partial / Due";

    // ===== UPDATE RESPONSE SHEET =====
    formSheet.getRange(row, 13).setValue(totalPaid);
    formSheet.getRange(row, 14).setValue(Math.max(dueAmount, 0));
    formSheet.getRange(row, 15).setValue(status);

    // ===== REGISTER SHEET =====
    var regData = registerSheet.getDataRange().getValues();
    var foundRow = -1;
    var existingPaid = 0;
    var existingCourseFee = courseFee;

    for (var i = 1; i < regData.length; i++) {
      if (String(regData[i][2]).trim() === phone) {
        foundRow = i + 1;
        existingCourseFee = Number(regData[i][5]) || courseFee; // col F
        existingPaid = Number(regData[i][6]) || 0;              // col G
        break;
      }
    }

    var newTotalPaid = 0;
    var newDue = 0;
    var newStatus = "";

    // ===== NEW STUDENT =====
    if (foundRow === -1) {

      var studentId = "JG-" + Utilities.formatDate(
        new Date(),
        "Asia/Dhaka",
        "yyyyMMddHHmmss"
      );

      newTotalPaid = paidNow;
      newDue = courseFee - newTotalPaid;
      newStatus = newDue <= 0 ? "Full Paid" : "Partial / Due";

      registerSheet.appendRow([
        studentId,
        name,
        phone,
        course,
        batch,
        courseFee,                 // Course Fee (fixed)
        newTotalPaid,
        Math.max(newDue, 0),
        newStatus
      ]);

      foundRow = registerSheet.getLastRow();

    } else {

      // ===== EXISTING STUDENT =====
      courseFee = existingCourseFee; // 🔥 IMPORTANT FIX

      newTotalPaid = existingPaid + paidNow;
      newDue = existingCourseFee - newTotalPaid;
      newStatus = newDue <= 0 ? "Full Paid" : "Partial / Due";

      registerSheet.getRange(foundRow, 7).setValue(newTotalPaid);
      registerSheet.getRange(foundRow, 8).setValue(Math.max(newDue, 0));
      registerSheet.getRange(foundRow, 9).setValue(newStatus);
    }

    // ===== FULL PAID HIGHLIGHT =====
    if (newStatus === "Full Paid") {
      registerSheet.getRange(foundRow, 1, 1, 9)
        .setBackground("#c6efce")
        .setFontColor("#006100");
    }

    // ===== LOGO & QR =====
    var logoBlob = DriveApp.getFileById(LOGO_ID).getBlob();
    var logoDataUrl = "data:image/png;base64," + Utilities.base64Encode(logoBlob.getBytes());

    var qrBlob = DriveApp.getFileById(QR_ID).getBlob();
    var qrDataUrl = "data:image/png;base64," + Utilities.base64Encode(qrBlob.getBytes());

    // ===== RECEIPT HTML =====
    var html = `
    <html>
    <body style="font-family:sans-serif; border:2px solid #ddd; padding:20px; max-width:500px; margin:auto;">
    
      <div style="text-align:center;">
        <img src="${logoDataUrl}" width="200"><br>
        <h2>JapanGate Language & Training Center</h2>
        <p>Dhaka, Bangladesh</p>
      </div>
      <hr>

      <p><b>Name:</b> ${name}</p>
      <p><b>Phone:</b> ${phone}</p>
      <p><b>Course:</b> ${course}</p>
      <p><b>Batch:</b> ${batch}</p>
      <p><b>Course Fee:</b> ${existingCourseFee || courseFee} BDT</p>
      <p><b>Paid Now:</b> ${paidNow} BDT</p>
      <p><b>Total Paid:</b> ${totalPaid} BDT</p>
      <p><b>Remaining:</b> ${Math.max(dueAmount,0)} BDT</p>
      <p><b>Status:</b> ${status}</p>
      <p><b>Date:</b> ${receiptDate}</p>

      <hr>

      <div style="text-align:center;">
        <img src="${qrDataUrl}" width="80"><br>
        <b>Thank You!</b>
      </div>

    </body>
    </html>
    `;

    // ===== PDF =====
    var blob = Utilities.newBlob(html, 'text/html').getAs('application/pdf');
    var pdfFile = DriveApp.getFolderById(MASTER_FOLDER_ID)
      .createFile(blob)
      .setName(`Receipt_${name}_${Date.now()}.pdf`);

    // ===== EMAIL =====
    MailApp.sendEmail({
      to: RECEIPT_EMAIL,
      subject: 'Payment Receipt - ' + name,
      htmlBody: 'Receipt attached.',
      attachments: [pdfFile]
    });

  } finally {
    lock.releaseLock();
  }
}