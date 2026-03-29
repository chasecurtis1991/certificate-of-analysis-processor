import { google } from 'googleapis';
import fs from 'fs';

const config = JSON.parse(fs.readFileSync('config.json', 'utf-8'));
const credentials = JSON.parse(fs.readFileSync(config.google.credentialsPath, 'utf-8'));
const { client_id, client_secret, redirect_uris } = credentials.installed;
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
const token = JSON.parse(fs.readFileSync(config.google.tokenPath, 'utf-8'));
oAuth2Client.setCredentials(token);
const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });

const sheetId = config.google.sheetId;
const sheetName = 'Sheet1';

// Test entries in alphabetical order
const testEntries = [
  { strain: 'Cosmic Taffy', category: 'Flower', driveId: '1Dt49fpJDxGnIT5cUFlUM5RwTlhv1Ye-p' },
  { strain: 'Grape Creamsicle', category: 'Flower', driveId: '1vyEFNaD1rqPhPS-ubamPBdavMcuoJmeP' },
  { strain: 'Gumbo', category: 'Flower', driveId: '1vyEFNaD1rqPhPS-ubamPBdavMcuoJmeP' },
  { strain: 'Mimosa', category: 'Flower', driveId: '1anwWiSvxvetSR1RGPMsZRwF8P8m5VGlX' },
];

async function addEntry(entry) {
  // Get current data
  const getRes = await sheets.spreadsheets.values.get({ spreadsheetId: sheetId, range: 'Sheet1!A:I' });
  const values = getRes.data.values || [];

  // Find table end (first empty row after header)
  let tableEnd = values.length;
  for (let i = 1; i < values.length; i++) {
    if (!values[i] || values[i].length === 0 || !values[i][0]) {
      tableEnd = i;
      break;
    }
  }

  // Find alphabetical insertion point
  let insertIdx = tableEnd;
  const strainLower = entry.strain.toLowerCase();
  for (let i = 1; i < tableEnd; i++) {
    const existingName = values[i]?.[0]?.toLowerCase() || '';
    if (strainLower < existingName) {
      insertIdx = i;
      break;
    }
  }

  const driveUrl = 'https://drive.google.com/file/d/' + entry.driveId + '/view';

  // Build row with HYPERLINK formulas (using comma for US locale)
  const row = [
    entry.strain,           // Name
    '',                     // Flower Type
    entry.category,         // Category
    '',                     // Report Date
    '',                     // Expiration Date
    'Yes',                  // Posted on Retail
    `=HYPERLINK("${driveUrl}","${entry.strain}")`,  // Google Drive COA Link (smart chip)
    `=HYPERLINK("${driveUrl}","${driveUrl}")`,      // Google Drive COA URL (plain hyperlink)
    'Auto-posted by COA Automation',  // Comments
  ];

  // Get sheet ID for batchUpdate
  const metaRes = await sheets.spreadsheets.get({ spreadsheetId: sheetId, fields: 'sheets.properties' });
  const googleSheetId = metaRes.data.sheets.find(s => s.properties.title === sheetName).properties.sheetId;

  // Insert row
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: sheetId,
    requestBody: {
      requests: [{
        insertDimension: {
          range: { sheetId: googleSheetId, dimension: 'ROWS', startIndex: insertIdx, endIndex: insertIdx + 1 },
        },
      }],
    },
  });

  // Populate the row
  const dataRowNum = insertIdx + 2;  // +1 for header, +1 for 1-indexing
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `Sheet1!A${dataRowNum}:I${dataRowNum}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [row] },
  });

  console.log(`Added ${entry.strain} at row ${dataRowNum} (insertIdx: ${insertIdx})`);
}

async function main() {
  console.log('Adding entries in alphabetical order...');

  for (const entry of testEntries) {
    await addEntry(entry);
  }

  // Verify the results
  console.log('\nVerifying results...');
  const getRes = await sheets.spreadsheets.values.get({ spreadsheetId: sheetId, range: 'Sheet1!A70:I85' });
  console.log('\nRows 70-85 after insert:');
  getRes.data.values.forEach((row, i) => console.log(70 + i, row));

  console.log('\nDone!');
}

main().catch(console.error);