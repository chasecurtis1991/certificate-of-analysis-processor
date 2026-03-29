/**
 * Pre-populate processed-files.json with all existing COA files in Drive
 * so the automation only processes truly new uploads.
 */
import { readFileSync, writeFileSync } from 'fs';
import { google } from 'googleapis';

const config = JSON.parse(readFileSync('./config.json', 'utf-8'));
const creds = JSON.parse(readFileSync(config.google.credentialsPath, 'utf-8'));
const { client_id, client_secret, redirect_uris } = creds.installed;
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
const token = JSON.parse(readFileSync(config.google.tokenPath, 'utf-8'));
oAuth2Client.setCredentials(token);

const drive = google.drive({ version: 'v3', auth: oAuth2Client });
const processed = {};
let total = 0;

for (const [type, folderId] of Object.entries(config.google.driveFolders)) {
  let pageToken = null;
  do {
    const res = await drive.files.list({
      q: `'${folderId}' in parents and mimeType='application/pdf' and trashed=false`,
      fields: 'nextPageToken, files(id, name)',
      pageSize: 100,
      pageToken,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });

    for (const file of res.data.files) {
      processed[file.id] = {
        name: file.name,
        strain: file.name.replace(/\.pdf$/i, '').replace(/_/g, ' ').replace(/\s*COA\s*(new)?\s*$/i, '').trim(),
        productType: type,
        skipped: 'pre-existing',
        processedAt: new Date().toISOString(),
      };
      total++;
    }

    pageToken = res.data.nextPageToken;
  } while (pageToken);

  console.log(`${type}: ${Object.values(processed).filter(p => p.productType === type).length} files`);
}

writeFileSync(config.processedFilesPath, JSON.stringify(processed, null, 2));
console.log(`\n✅ Seeded ${total} existing files into processed-files.json`);
