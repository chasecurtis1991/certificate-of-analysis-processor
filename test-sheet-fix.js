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

async function main() {
  // Get sheet ID
  const metaRes = await sheets.spreadsheets.get({ spreadsheetId: sheetId, fields: 'sheets.properties' });
  const googleSheetId = metaRes.data.sheets.find(s => s.properties.title === sheetName).properties.sheetId;

  // Delete rows 72-80 (the problematic rows)
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: sheetId,
    requestBody: {
      requests: [{
        deleteDimension: {
          range: { sheetId: googleSheetId, dimension: 'ROWS', startIndex: 71, endIndex: 80 },
        },
      }],
    },
  });

  console.log('Deleted rows 72-80');
}

main().catch(console.error);