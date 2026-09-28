const SHEET_NAME = 'Show Submissions'; // exact tab name
const DRIVE_FOLDER_INPUT = 'https://drive.google.com/drive/folders/1crfwbwRhl2zv85tc1QxU8o5LMXp3z6g9'; // your shared folder
const START_ROW = 2; // skip header row

// ---------- HEADER HELPERS ----------
// Looks up a column by its header text; creates the column (appended after
// the last used column) if it doesn't exist yet. Returns a 1-based column index.
// This makes the script resilient to Jotform re-ordering or adding columns
// when the form is duplicated/rebuilt.
function getOrCreateColumn_(sheet, headerName) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.indexOf(headerName);
  if (idx !== -1) return idx + 1;

  const newCol = lastCol + 1;
  sheet.getRange(1, newCol).setValue(headerName);
  return newCol;
}

function getColumnOrThrow_(headers, headerName) {
  const idx = headers.indexOf(headerName);
  if (idx === -1) throw new Error(`Missing expected column: "${headerName}"`);
  return idx; // 0-based, matches headers/data array indexing
}

// ---------- FULL ADDRESS ----------
function combineFullAddress() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const fullAddrCol = getOrCreateColumn_(sheet, 'Full Address');

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const streetCol = getColumnOrThrow_(headers, 'Show Address:');
  const cityCol   = getColumnOrThrow_(headers, 'TypeA28');
  const stateCol  = getColumnOrThrow_(headers, 'State');
  const fullAddrIdx = fullAddrCol - 1;

  for (let i = 1; i < data.length; i++) {
    const street = (data[i][streetCol] || '').toString().trim();
    const city   = (data[i][cityCol]   || '').toString().trim();
    const state  = (data[i][stateCol]  || '').toString().trim();
    const combined = [street, city, state].filter(Boolean).join(', ');
    if (combined && combined !== data[i][fullAddrIdx]) {
      sheet.getRange(i + 1, fullAddrCol).setValue(combined);
    }
  }
}

/**
 * Copies Jotform-uploaded images ("Upload a Picture" / "Upload a Second
 * Picture (Optional)") to Google Drive and writes public-view links into
 * "Photo 1" / "Photo 2" (created automatically if missing).
 */
function updateJotformImages() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  const photo1Col = getOrCreateColumn_(sheet, 'Photo 1');
  const photo2Col = getOrCreateColumn_(sheet, 'Photo 2');

  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const upload1Col = getColumnOrThrow_(headers, 'Upload a Picture') + 1;
  const upload2Col = getColumnOrThrow_(headers, 'Upload a Second Picture (Optional)') + 1;

  const folder = getFolderSafe_();
  const pairs = [
    { inputCol: upload1Col, outputCol: photo1Col },
    { inputCol: upload2Col, outputCol: photo2Col },
  ];

  let uploadedCount = 0;

  for (let row = START_ROW; row <= lastRow; row++) {
    pairs.forEach(({ inputCol, outputCol }) => {
      const url = sheet.getRange(row, inputCol).getValue();
      const output = sheet.getRange(row, outputCol).getValue();

      if (url && isJotformUrl_(url) && !output) {
        try {
          const blob = UrlFetchApp.fetch(url, {
            followRedirects: true,
            muteHttpExceptions: true,
          }).getBlob();

          const fileName = `row${row}_col${inputCol}.jpg`;
          const file = folder.createFile(blob).setName(fileName);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

          const publicUrl = `https://drive.google.com/uc?export=view&id=${file.getId()}`;
          sheet.getRange(row, outputCol).setValue(publicUrl);

          uploadedCount++;
        } catch (err) {
          Logger.log(`❌ Error row ${row}, col ${inputCol}: ${err}`);
        }
      }
    });
  }

  Logger.log(`✔️ Uploaded ${uploadedCount} new images to Drive.`);
}

// ---------- HELPERS ----------
function getFolderSafe_() {
  const idMatch = DRIVE_FOLDER_INPUT.match(/[-\w]{25,}/);
  if (!idMatch) throw new Error('Invalid folder ID or URL.');
  return DriveApp.getFolderById(idMatch[0]);
}
function isJotformUrl_(u) {
  return typeof u === 'string' && u.includes('jotform.com/uploads/');
}

// ---------- TEST ----------
function testFolderAccess() {
  const f = getFolderSafe_();
  Logger.log('✅ Folder OK: ' + f.getName() + ' (ID: ' + f.getId() + ')');
}

/**
 * Geocodes any row with a "Full Address" and fills in Latitude/Longitude
 * (columns created automatically if missing).
 */
function updateLatLon() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  const latCol = getOrCreateColumn_(sheet, 'Latitude');
  const lonCol = getOrCreateColumn_(sheet, 'Longitude');

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const addrIndex = getColumnOrThrow_(headers, 'Full Address');
  const latIndex = latCol - 1;
  const lonIndex = lonCol - 1;

  const geocoder = Maps.newGeocoder().setRegion('us');

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const address = row[addrIndex];
    const lat = row[latIndex];
    const lon = row[lonIndex];

    if (address && (lat === '' || lon === '' || lat === undefined)) {
      try {
        const resp = geocoder.geocode(address);
        if (resp.status === 'OK' && resp.results.length > 0) {
          const loc = resp.results[0].geometry.location;
          sheet.getRange(i + 1, latCol).setValue(loc.lat);
          sheet.getRange(i + 1, lonCol).setValue(loc.lng);
          Utilities.sleep(250);
        } else {
          Logger.log(`No result for ${address}`);
        }
      } catch (e) {
        Logger.log(`Error geocoding ${address}: ${e}`);
      }
    }
  }
}

function runAllOnSubmit() {
  combineFullAddress();
  updateJotformImages();
  updateLatLon();
}

// GET: manual/Home Assistant trigger. POST: Jotform's Webhooks integration
// (Settings > Integrations > Webhooks) posts here on every new submission.
function doGet(e) {
  return handleTrigger_();
}

function doPost(e) {
  return handleTrigger_();
}

function handleTrigger_() {
  Logger.log("✅ Trigger received.");
  runAllOnSubmit();
  SpreadsheetApp.flush();       // ensure changes are written

  // --- Get the sheet's real last-modified time from Drive ---
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(sheet.getId());
  const lastUpdatedISO = file.getLastUpdated().toISOString(); // 👉 ISO 8601 format

  // --- Build the response body ---
  const message =
    "✅ Light Up AZ map updated\n" +
    "Last Modified: " + lastUpdatedISO;

  // --- Return plain text ---
  return ContentService.createTextOutput(message)
    .setMimeType(ContentService.MimeType.TEXT);
}
