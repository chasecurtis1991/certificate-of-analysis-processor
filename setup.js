/**
 * One-time Google OAuth setup.
 * Run this first: node setup.js
 * It will print a URL — open it in your browser, authorize, then paste the code back.
 */

import { readFileSync, writeFileSync } from 'fs';
import { createInterface } from 'readline';
import { google } from 'googleapis';

const config = JSON.parse(readFileSync('./config.json', 'utf-8'));
const credentials = JSON.parse(readFileSync(config.google.credentialsPath, 'utf-8'));

const { client_id, client_secret, redirect_uris } = credentials.installed;
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

const SCOPES = [
  'https://www.googleapis.com/auth/drive.readonly',
  'https://www.googleapis.com/auth/spreadsheets',
];

const authUrl = oAuth2Client.generateAuthUrl({
  access_type: 'offline',
  scope: SCOPES,
  prompt: 'consent',
});

console.log('\n=== Google OAuth Setup ===\n');
console.log('1. Open this URL in your browser:\n');
console.log(authUrl);
console.log('\n2. Authorize the application');
console.log('3. Copy the authorization code from the redirect URL');
console.log('   (It will redirect to localhost — grab the "code" parameter from the URL)\n');

const rl = createInterface({ input: process.stdin, output: process.stdout });

rl.question('Paste the authorization code here: ', async (code) => {
  rl.close();
  try {
    const { tokens } = await oAuth2Client.getToken(code);
    writeFileSync(config.google.tokenPath, JSON.stringify(tokens, null, 2));
    console.log('\n✅ Token saved to', config.google.tokenPath);
    console.log('You can now run: npm start');
  } catch (err) {
    console.error('\n❌ Error getting token:', err.message);
  }
});
