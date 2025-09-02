// === CONFIG ===
const SHEET_NAME = "Signups";  
const SPREADSHEET_ID = "1N-WM0XfWwL2cJ9zxnoTS2FhdaLq5p46y8Z9ofwNqJuw";  
const WEBAPP_URL = "https://script.google.com/macros/s/AKfycbyMDFstF2qj9Lep6j06WGXnxNpgSzTc-Y5tRgsjJXhQZ8qRj6nNX6qAr9BJUHm3Lwks/exec";

// === HOMEPAGE HANDLER ===
function onHomepage(e) {
  const userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    return createResponseCard("❌ Tidak bisa ambil email Google. Pastikan kamu login.");
  }

  let activeDocId, activeDocName;
  try {
    const activeDoc = DocumentApp.getActiveDocument();
    activeDocId = activeDoc.getId();
    activeDocName = activeDoc.getName();
  } catch (err) {
    activeDocId = null;
    activeDocName = null;
  }

  const docs = getUserDocs(userEmail);

  // 1️⃣ Match dengan Doc ID dari Spreadsheet
  let match = null;
  if (activeDocId) {
    match = docs.find(doc => doc.docId === activeDocId);
  }

  // 2️⃣ Kalau belum ketemu, coba match berdasarkan nama file
  if (!match && activeDocName) {
    match = docs.find(doc => doc.name === activeDocName);
  }

  // ✅ Kalau ada match di Spreadsheet → tampilkan tombol sign/print QR
  if (match) {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle("Ivosights eSignature"))
      .addSection(
        CardService.newCardSection()
          .addWidget(CardService.newTextParagraph().setText(`📄 Dokumen terdeteksi: ${match.name}`))
          .addWidget(
            CardService.newTextButton()
              .setText("✍️ Tandatangani & Cetak QR")
              .setOnClickAction(
                CardService.newAction()
                  .setFunctionName("signDocumentFromSheet")
                  .setParameters({ recordId: match.id })
              )
          )
      )
      .build();
  }

  // 3️⃣ Kalau tidak ada match di Spreadsheet → ambil dari isi Docs
  const docIdValue = getDocumentIDValue();
  const projectName = getProjectNameValue();
  const combinedName = getCombinedDocName();

  if (docIdValue || projectName) {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle("Ivosights eSignature"))
      .addSection(
        CardService.newCardSection()
          .addWidget(
            CardService.newTextParagraph().setText(
              `📄 Dokumen aktif tidak ada di Spreadsheet,<br>tapi berhasil dibaca dari isi file:<br><br>
               <b>Document ID:</b> ${docIdValue || "-"}<br>
               <b>Project Name:</b> ${projectName || "-"}<br>
               <b>Combined:</b> ${combinedName}`
            )
          )
          .addWidget(
            CardService.newTextButton()
              .setText("✍️ Tandatangani & Cetak QR")
              .setOnClickAction(
                CardService.newAction()
                  .setFunctionName("printQrFromDocs")
                  .setParameters({
                    docIdValue: docIdValue || "",
                    projectName: projectName || "",
                    combinedName: combinedName || ""
                  })
              )
          )
      )
      .build();
  }

  // fallback terakhir kalau semua gagal
  return createResponseCard("⚠️ Dokumen aktif tidak terdaftar untuk akun kamu.");
}

// === HELPER UNTUK RESPONSE CARD ===
function createResponseCard(msg) {
  return CardService.newCardBuilder()
    .addSection(
      CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText(msg))
    )
    .build();
}

// === Ambil daftar dokumen dari Spreadsheet ===
function getUserDocs(userEmail) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  const data = sh.getDataRange().getValues();
  const rows = data.slice(1);

  return rows
    .filter(r => r[3] === userEmail) // kolom email
    .map(r => ({
      id: r[0],        // recordId
      timestamp: r[1],
      name: r[6],      // docName
      docId: r[5],     // docId
      docUrl: r[10] || "" // kolom URL dokumen
    }));
}

// === Sign dokumen berdasarkan record di Spreadsheet ===
function signDocumentFromSheet(e) {
  const recordId = e.parameters.recordId;
  if (!recordId) {
    return createResponseCard("❌ Record ID tidak ditemukan.");
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  const data = sh.getDataRange().getValues();
  const rows = data.slice(1);
  const record = rows.find(r => r[0] === recordId);

  if (!record) {
    return createResponseCard("❌ Data tidak ditemukan di Spreadsheet.");
  }

  const docId = record[5];
  const docName = record[6];
  const userEmail = record[3];
  const signByName = userEmail.split("@")[0];
  const signTime = new Date();

  const doc = DocumentApp.openById(docId);

  // === Buat QR ===
  const detailUrl = `${WEBAPP_URL}?id=${recordId}`;
  const qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" + encodeURIComponent(detailUrl);
  const qrBlob = UrlFetchApp.fetch(qrUrl).getBlob().setName(`qr_${recordId}.png`);

  // Insert QR
  doc.getBody().appendImage(qrBlob).setWidth(96).setHeight(96);
  doc.saveAndClose();

  // Notifikasi Email
  GmailApp.sendEmail(
    userEmail,
    `Ivosights eSignature: ${docName} signed`,
    "See details in HTML body",
    {
      htmlBody: `
        <p>Dokumen <b>${docName}</b> telah ditandatangani oleh ${signByName}.</p>
        <p><a href="${doc.getUrl()}">Buka Dokumen</a></p>
        <p><a href="${detailUrl}">Detail Signature</a></p>
        <p><img src="cid:qr"></p>
      `,
      inlineImages: { qr: qrBlob }
    }
  );

  return createResponseCard("✅ Dokumen berhasil ditandatangani!");
}

// === Ambil Document ID dari tabel di Docs ===
function getDocumentIDValue() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var tables = body.getTables();
  
  var docIdValue = null;
  
  for (var t = 0; t < tables.length; t++) {
    var table = tables[t];
    for (var r = 0; r < table.getNumRows(); r++) {
      var row = table.getRow(r);
      var firstCellText = row.getCell(0).getText().trim();
      
      if (firstCellText === "Document ID") {
        docIdValue = row.getCell(1).getText().trim();
        break;
      }
    }
    if (docIdValue) break;
  }
  
  return docIdValue;
}

// === Ambil Project Name dari tabel di Docs ===
function getProjectNameValue() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var tables = body.getTables();
  
  var projectName = null;
  
  for (var t = 0; t < tables.length; t++) {
    var table = tables[t];
    for (var r = 0; r < table.getNumRows(); r++) {
      var row = table.getRow(r);
      var firstCellText = row.getCell(0).getText().trim();
      
      if (firstCellText === "Project Name") {
        projectName = row.getCell(1).getText().trim();
        break;
      }
    }
    if (projectName) break;
  }
  
  return projectName;
}

// === Gabungkan nama file Google Docs dengan Project Name ===
function getCombinedDocName() {
  var doc = DocumentApp.getActiveDocument();
  var docName = doc.getName();
  var projectName = getProjectNameValue();
  
  if (projectName) {
    return docName + " - " + projectName;
  }
  return docName;
}

// === Print QR dari Spreadsheet Signups sesuai Document ID di Docs ===
function printQrFromDocs(e) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  const data = sh.getDataRange().getValues();
  const rows = data.slice(1);

  // Ambil Document ID dari isi Docs
  const docIdValue = getDocumentIDValue();
  if (!docIdValue) {
    DocumentApp.getUi().alert("❌ Document ID tidak ditemukan di tabel Docs.");
    return;
  }

  // Cari record berdasarkan docId di Spreadsheet
  const record = rows.find(r => r[5] === docIdValue); // kolom [5] = docId
  if (!record) {
    DocumentApp.getUi().alert("❌ Document ID tidak terdaftar di Spreadsheet.");
    return;
  }

  const recordId = record[0];
  const docName = record[6];
  const userEmail = record[3];

  // === Buat QR berdasarkan recordId ===
  const detailUrl = `${WEBAPP_URL}?id=${recordId}`;
  const qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" + encodeURIComponent(detailUrl);
  const qrBlob = UrlFetchApp.fetch(qrUrl).getBlob().setName(`qr_${recordId}.png`);

  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();
  const body = doc.getBody();

  // Tempel QR di posisi cursor atau akhir dokumen
  if (cursor) {
    const img = cursor.insertInlineImage(qrBlob);
    img.setWidth(96).setHeight(96);
  } else {
    body.appendParagraph("\n--- QR Code ---");
    body.appendImage(qrBlob).setWidth(96).setHeight(96);
  }

  doc.saveAndClose();

  Logger.log(`✅ QR dari Spreadsheet ditempel di dokumen: ${docName}`);
}





