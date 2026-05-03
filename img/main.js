
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
    var classStart  = data[5];
    var courseFee   = Number(data[6]) || 0;
    var paidNow     = Number(data[7]) || 0;
    var payBy       = data[8];
    var authorized  = data[9];
    var comment     = data[10];
    var dueDateRaw  = data[11];

    var receiptDate = Utilities.formatDate(
      new Date(timestamp),
      Session.getScriptTimeZone(),
      "dd MMM yyyy"
    );

    var dueDate = dueDateRaw
      ? Utilities.formatDate(new Date(dueDateRaw), Session.getScriptTimeZone(), "dd MMM yyyy")
      : "";

    // ===== TOTAL PAID CALCULATION (PHONE BASED) =====
    var lastRow = formSheet.getLastRow();
    var allData = formSheet.getRange(2, 1, lastRow - 1, 8).getValues();

    var totalPaid = 0;
    var originalCourseFee = courseFee;

    allData.forEach(function(r) {
      if (String(r[2]).trim() === phone) {
        totalPaid += Number(r[7]) || 0;
        originalCourseFee = Number(r[6]) || originalCourseFee;
      }
    });

    var dueAmount = originalCourseFee - totalPaid;
    var status = dueAmount <= 0 ? "Full Paid" : "Partial / Due";

    // ===== UPDATE SAME ROW =====
    formSheet.getRange(row, 13).setValue(totalPaid);              // M → Total Paid
    formSheet.getRange(row, 14).setValue(Math.max(dueAmount, 0)); // N → Remaining
    formSheet.getRange(row, 15).setValue(status);                 // O → Status

   
    // ===== UPDATE REGISTER SHEET =====
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

    // ===== IF NEW STUDENT =====
    if (foundRow === -1) {

      var studentId = "JG-" + Utilities.formatDate(
        new Date(),
        "Asia/Dhaka",
        "yyyyMMddHHmmss"
      );

      var newTotalPaid = paidNow;
      var newDue = courseFee - newTotalPaid;
      var newStatus = newDue <= 0 ? "Full Paid" : "Partial / Due";

      // registerSheet.appendRow([
      //   studentId,
      //   name,
      //   phone,
      //   course,
      //   batch,
      //   newTotalPaid,
      //   Math.max(newDue, 0),
      //   newStatus,
      //   "No"
      // ]);

      registerSheet.appendRow([
      studentId,
      name,
      phone,
      course,
      batch,
      courseFee,          // ✅ NEW
      paidNow,            // total paid শুরু
      courseFee - paidNow,
      (courseFee - paidNow <= 0 ? "Full Paid" : "Partial / Due")
    ]);


      foundRow = registerSheet.getLastRow();

    } else {

  // ✅ IMPORTANT: course fee override
  courseFee = existingCourseFee;

  // ===== EXISTING STUDENT UPDATE =====
  var newTotalPaid = existingPaid + paidNow;
  var newDue = existingCourseFee - newTotalPaid;
  var newStatus = newDue <= 0 ? "Full Paid" : "Partial / Due";

  registerSheet.getRange(foundRow, 7).setValue(newTotalPaid);
  registerSheet.getRange(foundRow, 8).setValue(Math.max(newDue, 0));
  registerSheet.getRange(foundRow, 9).setValue(newStatus);
}
    // ===== FULL PAID হলে GREEN =====
    if (newStatus === "Full Paid") {
      registerSheet.getRange(foundRow, 1, 1, 9)
        .setBackground("#c6efce")
        .setFontColor("#006100");
    }

    // ===== LOGO & QR =====
    var logoBlob = DriveApp.getFileById(LOGO_ID).getBlob();
    var logoBase64 = Utilities.base64Encode(logoBlob.getBytes());
    var logoDataUrl = "data:image/png;base64," + logoBase64;

    var qrBlob = DriveApp.getFileById(QR_ID).getBlob();
    var qrBase64 = Utilities.base64Encode(qrBlob.getBytes());
    var qrDataUrl = "data:image/png;base64," + qrBase64;

    // ===== RECEIPT HTML =====
    var html = `
    <html>
    <body style="font-family:sans-serif; border:2px solid #ddd; padding:20px; max-width:500px; margin:auto;">
    
      <div style="text-align:center;">
        <img src="${logoDataUrl}" width="200"><br>
        <h2>JapanGate Language & Training Center</h2>
        <p>House No -128, China Building Goli, Azimpur Road 1205 Dhaka, Bangladesh</p>
        <p style="font-size:12px;">📞 01577-481249 | 📧 contact.japangatebd@gmail.com</p>
      </div>
      <hr>

      <p><b>Name:</b> ${name}</p>
      <p><b>Phone:</b> ${phone}</p>
      <p><b>Course:</b> ${course}</p>
      <p><b>Batch:</b> ${batch}</p>
      <p><b>Course Fee:</b> ${originalCourseFee} BDT</p>
      <p><b>Paid Now:</b> ${paidNow} BDT</p>
      <p><b>Total Paid:</b> ${totalPaid} BDT</p>
      <p><b>Remaining:</b> ${Math.max(dueAmount,0)} BDT</p>
      <p><b>Status:</b> ${status}</p>
      <p><b>Date:</b> ${receiptDate}</p>

      <hr>

      <div style="text-align:center;">
        <img src="${qrDataUrl}" width="80"><br>
        <b>Thank You for choosing JapanGate!</b>
      </div>

    </body>
    </html>
    `;

    // ===== PDF CREATE =====
    var blob = Utilities.newBlob(html, 'text/html').getAs('application/pdf');
    var masterFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    var fileName = `Receipt_${name}_${Date.now()}.pdf`;
    var pdfFile = masterFolder.createFile(blob).setName(fileName);

    // ===== EMAIL =====
    MailApp.sendEmail({
      to: 'bijoyhaque8@gmail.com, bageshut@gmail.com',
      subject: 'New Payment Receipt - ' + name,
      htmlBody: 'Receipt generated.<br><a href="' + pdfFile.getUrl() + '">Open Receipt</a>',
      attachments: [pdfFile]
    });

  } finally {
    lock.releaseLock();
  }
}


