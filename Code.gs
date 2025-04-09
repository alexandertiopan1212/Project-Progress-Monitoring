function markKeyMilestones() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master");
  const startRow = 7;
  const lastRow = sheet.getLastRow();
  const noCol = 2; // Kolom B (No)
  const finishCol = 11; // Kolom K (Actual Finish)
  const keyMilestoneCol = 12; // Kolom L (Key Milestone)

  let currentParent = null;
  let childrenMap = {};

  // Langkah 1: Identifikasi struktur parent-child
  for (let row = startRow; row <= lastRow; row++) {
    const noVal = sheet.getRange(row, noCol).getValue();
    if (noVal) {
      currentParent = row;
      childrenMap[currentParent] = [];
    } else if (currentParent) {
      childrenMap[currentParent].push(row);
    }
  }

  // Bersihkan isi kolom Key Milestone terlebih dahulu
  sheet.getRange(startRow, keyMilestoneCol, lastRow - startRow + 1).clearContent();

  // Langkah 2: Tandai task yang merupakan bagian dari jalur kritis
  for (let parentRow in childrenMap) {
    const children = childrenMap[parentRow];
    if (children.length === 0) continue;

    let maxFinishDate = null;
    let criticalChildRow = null;

    for (let row of children) {
      const finishDate = sheet.getRange(row, finishCol).getValue();
      if (finishDate instanceof Date) {
        if (!maxFinishDate || finishDate > maxFinishDate) {
          maxFinishDate = finishDate;
          criticalChildRow = row;
        }
      }
    }

    if (criticalChildRow && maxFinishDate) {
      const parentFinishDate = sheet.getRange(Number(parentRow), finishCol).getValue();
      if (parentFinishDate instanceof Date && parentFinishDate.getTime() === maxFinishDate.getTime()) {
        sheet.getRange(Number(parentRow), keyMilestoneCol).setValue("CRITICAL");
        sheet.getRange(criticalChildRow, keyMilestoneCol).setValue("CRITICAL");
      }
    }
  }
}



function generateGantt() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master");
  const startRow = 7;
  const noCol = 2; // Kolom B
  const planningStartCol = 8;  // H
  const planningFinishCol = 9; // I
  const actualStartCol = 10;   // J
  const actualFinishCol = 11;  // K
  const ganttStartCol = 15;    // Kolom O
  const totalWeeks = 20;
  const ganttStartDate = new Date("2025-04-01");
  const lastRow = sheet.getLastRow();

  // ✅ Normalize ke hari Senin pertama
  const normalizedStart = new Date(ganttStartDate);
  const day = normalizedStart.getDay(); // 0 = Minggu, 1 = Senin, ..., 6 = Sabtu
  const diffToMonday = (day + 6) % 7;
  normalizedStart.setDate(normalizedStart.getDate() - diffToMonday);

  let childrenMap = {};
  let currentParent = null;

  // 🧭 Identifikasi parent-child
  for (let row = startRow; row <= lastRow; row++) {
    const noVal = sheet.getRange(row, noCol).getValue();
    if (noVal) {
      currentParent = row;
      childrenMap[currentParent] = [];
    } else if (currentParent) {
      childrenMap[currentParent].push(row);
    }
  }

  for (let row = startRow; row <= lastRow; row++) {
    let planningStart, planningEnd, actualStart, actualEnd;

    if (row in childrenMap && childrenMap[row].length > 0) {
      const childRows = childrenMap[row];
      planningStart = getMinDate(sheet, childRows, planningStartCol);
      planningEnd = getMaxDate(sheet, childRows, planningFinishCol);
      actualStart = getMinDate(sheet, childRows, actualStartCol);
      actualEnd = getMaxDate(sheet, childRows, actualFinishCol);

      if (planningStart) sheet.getRange(row, planningStartCol).setValue(planningStart);
      if (planningEnd) sheet.getRange(row, planningFinishCol).setValue(planningEnd);
      if (actualStart) sheet.getRange(row, actualStartCol).setValue(actualStart);
      if (actualEnd) sheet.getRange(row, actualFinishCol).setValue(actualEnd);
    } else {
      planningStart = sheet.getRange(row, planningStartCol).getValue();
      planningEnd = sheet.getRange(row, planningFinishCol).getValue();
      actualStart = sheet.getRange(row, actualStartCol).getValue();
      actualEnd = sheet.getRange(row, actualFinishCol).getValue();
    }

    // 🎨 Gambar Gantt Chart Mingguan
    for (let i = 0; i < totalWeeks; i++) {
      const weekStart = new Date(normalizedStart);
      weekStart.setDate(normalizedStart.getDate() + i * 7);
      const weekEnd = new Date(weekStart);
      weekEnd.setDate(weekStart.getDate() + 6);

      // const cell = sheet.getRange(row, ganttStartCol + i);
      // cell.setBackground("#ffffff").setBorder(true, true, true, true, null, null, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);

      const cell = sheet.getRange(row, ganttStartCol + i);
      cell.setBackground("#FFFFFF")
          .setBorder(true, true, true, true, null, null, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);


      const isPlanning = isDateInWeek(planningStart, planningEnd, weekStart, weekEnd);
      const isActual = isDateInWeek(actualStart, actualEnd, weekStart, weekEnd);

//       const keyMilestoneVal = sheet.getRange(row, 12).getValue(); // kolom L (Key Milestone)

//       let color = "#FFFFFF"; // default color

//       if (keyMilestoneVal === "CRITICAL") {
//         if (isPlanning && isActual) {
//           color = "#DFA7E8"; // Critical overlap (soft purple)
//         } else if (isPlanning) {
//           color = "#9DC3E6"; // Critical Planning (soft blue)
//         } else if (isActual) {
//           color = "#A9D18E"; // Critical Actual (soft green)
//         }
//       } else {
//         if (isPlanning && isActual) {
//           color = "#B4E1E1"; // Overlap biasa (soft teal)
//         } else if (isPlanning) {
//           color = "#2F5597"; // Planning biasa (darker blue)
//         } else if (isActual) {
//           color = "#548235"; // Actual biasa (darker green)
//         }
// }

        // cell.setBackground(color);

      if (isPlanning && isActual) {
        cell.setBackground("#008080"); // 💜 Overlap
      } else if (isPlanning) {
        cell.setBackground("#2F5597"); // 🔵 Planning
      } else if (isActual) {
        cell.setBackground("#548235"); // 🔴 Actual
      }
    }
  }
}

// ✅ Periksa apakah durasi aktivitas jatuh di dalam minggu tertentu
function isDateInWeek(start, end, weekStart, weekEnd) {
  if (!(start instanceof Date) || !(end instanceof Date)) return false;

  const s = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const e = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  const ws = new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate());
  const we = new Date(weekEnd.getFullYear(), weekEnd.getMonth(), weekEnd.getDate());

  return e >= ws && s <= we;
}

// 🔧 Ambil tanggal paling awal dari anak
function getMinDate(sheet, rows, col) {
  let minDate = null;
  for (let row of rows) {
    const val = sheet.getRange(row, col).getValue();
    if (val instanceof Date && (!minDate || val < minDate)) {
      minDate = val;
    }
  }
  return minDate;
}

// 🔧 Ambil tanggal paling akhir dari anak
function getMaxDate(sheet, rows, col) {
  let maxDate = null;
  for (let row of rows) {
    const val = sheet.getRange(row, col).getValue();
    if (val instanceof Date && (!maxDate || val > maxDate)) {
      maxDate = val;
    }
  }
  return maxDate;
}

// 🔁 Auto Generate saat buka
function onOpen() {
  generateGantt();
  markKeyMilestones(); // ✅ Auto detect critical path
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const col = e.range.getColumn();
  if (sheet.getName() === "Master" && col >= 8 && col <= 11) {
    generateGantt();
    markKeyMilestones(); // Tambahkan di sini juga kalau mau update real-time
  }
}

// 🔁 Trigger otomatis
function createTimeTrigger() {
  ScriptApp.newTrigger("generateGantt")
    .timeBased()
    .everyMinutes(5)
    .create();
}

function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
}



function createDailyReport() {
  const html = HtmlService.createHtmlOutputFromFile('DatePickerDialog')
    .setWidth(300)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'Pilih Tanggal Daily Report');
}

function createSheetAndPickActivities(dateString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formattedDate = Utilities.formatDate(new Date(dateString), Session.getScriptTimeZone(), "yyyy-MM-dd");

  if (ss.getSheetByName(formattedDate)) {
    return { status: "exists", sheetName: formattedDate };
  }

  const templateFileId = "1inYDICjfaovcVzZBm6JNT02NfM_MELUZnh0cokXFpqw"; // ganti ID kalau perlu
  const templateSheetName = "Template_Daily_Report";
  const templateFile = SpreadsheetApp.openById(templateFileId);
  const templateSheet = templateFile.getSheetByName(templateSheetName);

  if (!templateSheet) {
    return { status: "error", message: "Template sheet tidak ditemukan." };
  }

  const newSheet = templateSheet.copyTo(ss).setName(formattedDate);
  ss.setActiveSheet(newSheet);
  PropertiesService.getScriptProperties().setProperty("activeDailySheet", formattedDate);

  return {
    status: "success",
    activities: getActivityListFromMaster()
  };
}

function getActivityListFromMaster() {
  const master = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master");
  const data = master.getRange("B7:C").getValues(); // B = Parent, C = Child

  const activityList = [];
  let currentParent = null;
  const parentWithChild = new Set();

  // Step 1: Identify parent yang punya child
  data.forEach(([parent, child]) => {
    if (parent && child) {
      parentWithChild.add(parent.trim());
    }
  });

  // Step 2: Build final activity list
  data.forEach(([parent, child]) => {
    if (parent) {
      currentParent = parent.trim();
    }

    if (child && currentParent) {
      activityList.push(`${currentParent} - ${child.trim()}`);
    } else if (!child && parent && !parentWithChild.has(parent.trim())) {
      // parent tanpa child
      activityList.push(parent.trim());
    }
  });

  return activityList;
}


function showActivitySelector() {
  const html = HtmlService.createHtmlOutputFromFile("ActivitySelector")
    .setWidth(400)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, "Pilih Aktivitas Hari Ini");
}

function insertSelectedActivitiesToSheet(selectedActivities) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = PropertiesService.getScriptProperties().getProperty("activeDailySheet");
  const sheet = ss.getSheetByName(sheetName);
  const col = 2; // Kolom "Activities"

  // Cari row tempat header "Activities"
  const values = sheet.getRange("B1:B").getValues();
  let startRow = -1;
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] && values[i][0].toString().trim().toLowerCase() === "activities") {
      startRow = i + 2; // baris setelah header
      break;
    }
  }

  if (startRow === -1) {
    throw new Error('Header "Activities" tidak ditemukan di kolom B');
  }

  const existingValues = sheet.getRange(startRow, col, sheet.getLastRow() - startRow + 1).getValues().flat();
  const messages = [];
  let insertRow = startRow;

  selectedActivities.forEach(act => {
    if (existingValues.includes(act)) {
      messages.push(`❌ "${act}" sudah ada di sheet.`);
    } else {
      // Cari baris kosong berikutnya
      while (sheet.getRange(insertRow, col).getValue()) {
        insertRow++;
      }
      sheet.getRange(insertRow, col).setValue(act);
      messages.push(`✅ "${act}" berhasil ditambahkan.`);
    }
  });

  return messages.join('\n');
}


// Buka popup untuk pilih next activities
function insertNextDayActivities() {
  const html = HtmlService.createHtmlOutputFromFile('NextDaySelector')
    .setWidth(400)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, "Pilih Next Day Activities");
}

// Ambil daftar aktivitas dari Master (Parent + Child)
function getNextDayActivityListFromMaster() {
  const master = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master");
  const data = master.getRange("B7:C").getValues(); // B = Parent, C = Child

  const activityList = [];
  let currentParent = null;
  const parentWithChild = new Set();

  // Step 1: Tandai parent yang punya child
  data.forEach(([parent, child]) => {
    if (parent && child) {
      parentWithChild.add(parent.trim());
    }
  });

  // Step 2: Build list aktivitas
  data.forEach(([parent, child]) => {
    if (parent) currentParent = parent.trim();

    if (child && currentParent) {
      activityList.push(`${currentParent} - ${child.trim()}`);
    } else if (parent && !child && !parentWithChild.has(parent.trim())) {
      activityList.push(parent.trim());
    }
  });

  return activityList;
}

// Masukkan aktivitas yang dipilih ke bagian "Next Activities"
function insertNextSelectedActivitiesToSheet(selectedActivities) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = PropertiesService.getScriptProperties().getProperty("activeDailySheet");
  const sheet = ss.getSheetByName(sheetName);
  const col = 2; // Kolom B

  // Cari baris header "Next Activities"
  const values = sheet.getRange("B1:B" + sheet.getMaxRows()).getValues();
  let startRow = -1;
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] && values[i][0].toString().trim().toLowerCase() === "next activities") {
      startRow = i + 2; // satu baris di bawah header
      break;
    }
  }

  if (startRow === -1) {
    SpreadsheetApp.getUi().alert('❌ Header "Next Activities" tidak ditemukan!');
    return;
  }

  // AMAN: ambil 100 baris ke bawah dari startRow (bisa disesuaikan kalau mau)
  const checkRows = 100;
  const existingRange = sheet.getRange(startRow, col, checkRows);
  const existingValues = existingRange.getValues().flat().map(x => x?.toString()?.trim());

  const messages = [];
  let insertRow = startRow;

  selectedActivities.forEach(act => {
    if (existingValues.includes(act)) {
      messages.push(`❌ "${act}" sudah ada di bagian Next Day.`);
    } else {
      while (sheet.getRange(insertRow, col).getValue()) {
        insertRow++;
      }
      sheet.getRange(insertRow, col).setValue(act);
      messages.push(`✅ "${act}" ditambahkan ke Next Day.`);
    }
  });

  return messages.join('\n');
}


function showIssuePhotoUploader() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Issues");
  const cell = sheet.getActiveCell();

  if (cell.getColumn() !== 8) { // Kolom "Photos"
    SpreadsheetApp.getUi().alert("📸 Pilih dulu cell di kolom 'Photos'.");
    return;
  }

  const row = cell.getRow();
  PropertiesService.getScriptProperties().setProperty("targetRow", row.toString());

  const html = HtmlService.createHtmlOutputFromFile('IssuePhotoUploader')
    .setTitle("Upload Foto Issue");
  SpreadsheetApp.getUi().showSidebar(html);
}

function getTargetRow() {
  return PropertiesService.getScriptProperties().getProperty("targetRow");
}

function uploadIssuePhoto(dataUrl, filename) {
  const folderId = "1VP8qUTwV7HZW_Q22udCOrryuBm_nKS-S"; // Ganti dengan folder Drive kamu
  const folder = DriveApp.getFolderById(folderId);

  // Ambil tipe file dan base64-nya
  const contentType = dataUrl.match(/^data:(image\/\w+);base64,/)[1];
  const base64Data = dataUrl.replace(/^data:image\/\w+;base64,/, "");
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, filename);

  const file = folder.createFile(blob);
  const url = file.getUrl();

  // Masukkan link ke sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Issues");
  const row = parseInt(PropertiesService.getScriptProperties().getProperty("targetRow"));
  sheet.getRange(row, 8).setFormula(`=HYPERLINK("${url}", "View")`);
}
