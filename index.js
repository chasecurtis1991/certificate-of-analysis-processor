/**
 * COA Automation
 *
 * Watches Google Drive for new COA PDFs (and images), uploads them to Shopify,
 * updates the storefront theme template, and logs entries to Google Sheets.
 *
 * Usage:
 *   node index.js          - Continuous polling
 *   node index.js --once   - Single run then exit
 */

import { readFileSync, writeFileSync, existsSync } from 'fs';
import { google } from 'googleapis';
import pdfParse from 'pdf-parse';

// ── Config ──────────────────────────────────────────────────────────────────

const config = JSON.parse(readFileSync('./config.json', 'utf-8'));
const ONCE_MODE = process.argv.includes('--once');

// ── Logging ─────────────────────────────────────────────────────────────────

function log(msg) {
  console.log(`[${new Date().toISOString()}] ${msg}`);
}

function logError(msg, err) {
  console.error(`[${new Date().toISOString()}] ❌ ${msg}`, err?.message || err);
}

// ── Google Auth ─────────────────────────────────────────────────────────────

function getGoogleAuth() {
  const credentials = JSON.parse(readFileSync(config.google.credentialsPath, 'utf-8'));
  const { client_id, client_secret, redirect_uris } = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

  if (!existsSync(config.google.tokenPath)) {
    throw new Error('No token.json found. Run "npm run setup" first.');
  }

  const token = JSON.parse(readFileSync(config.google.tokenPath, 'utf-8'));
  oAuth2Client.setCredentials(token);

  // Auto-refresh and save new tokens
  oAuth2Client.on('tokens', (newTokens) => {
    const existing = JSON.parse(readFileSync(config.google.tokenPath, 'utf-8'));
    const merged = { ...existing, ...newTokens };
    writeFileSync(config.google.tokenPath, JSON.stringify(merged, null, 2));
    log('Google token refreshed and saved.');
  });

  return oAuth2Client;
}

// ── Processed Files Tracking ────────────────────────────────────────────────

function loadProcessedFiles() {
  if (existsSync(config.processedFilesPath)) {
    return JSON.parse(readFileSync(config.processedFilesPath, 'utf-8'));
  }
  return {};
}

function saveProcessedFiles(processed) {
  writeFileSync(config.processedFilesPath, JSON.stringify(processed, null, 2));
}

// ── Google Drive ────────────────────────────────────────────────────────────

async function listNewFiles(drive, processed) {
  const allNewFiles = [];

  for (const [type, folderId] of Object.entries(config.google.driveFolders)) {
    // NOTE: Originally this automation only processed COA PDFs.
    // Some vendors/users upload image files (e.g., PNG) as COAs; allow those too.
    const res = await drive.files.list({
      q: `'${folderId}' in parents and (mimeType='application/pdf' or mimeType='image/png' or mimeType='image/jpeg') and trashed=false`,
      fields: 'files(id, name, mimeType, webViewLink, webContentLink, createdTime)',
      orderBy: 'createdTime desc',
      pageSize: 100,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });

    const files = (res.data.files || [])
      .filter(f => !processed[f.id])
      .map(f => ({ ...f, productType: type }));

    allNewFiles.push(...files);
  }

  return allNewFiles;
}

async function downloadFile(drive, fileId) {
  const res = await drive.files.get(
    { fileId, alt: 'media', supportsAllDrives: true },
    { responseType: 'arraybuffer' }
  );
  return Buffer.from(res.data);
}

// ── PDF Date Extraction ─────────────────────────────────────────────────────

async function extractReportDate(pdfBuffer) {
  try {
    const data = await pdfParse(pdfBuffer);
    const text = data.text;

    // Strategy: Look for date patterns near common labels
    const labels = [
      'completed',
      'report created',
      'report date',
      'date reported',
      'date completed',
      'date of report',
      'signed on',
      'analysis date',
      'date issued',
      'issued',
    ];

    // Helper: normalize 2-digit year to 4-digit
    function normalizeYear(y) {
      const yr = parseInt(y);
      if (yr < 100) return yr + 2000;
      return yr;
    }

    // Helper: format date as MM/DD/YYYY
    function formatDate(month, day, year) {
      return `${String(month).padStart(2, '0')}/${String(day).padStart(2, '0')}/${year}`;
    }

    // Helper: parse "Month DD, YYYY" format
    const monthNames = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

    // Pass 1: Look for "Month DD, YYYY" near labels (most reliable for Labstat)
    for (const label of labels) {
      const labelIdx = text.toLowerCase().indexOf(label);
      if (labelIdx === -1) continue;
      const snippet = text.substring(labelIdx, labelIdx + 200);
      const monthMatch = snippet.match(/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+(\d{1,2}),?\s+(\d{4})/i);
      if (monthMatch) {
        const month = monthNames[monthMatch[1].toLowerCase().substring(0, 3)];
        const day = parseInt(monthMatch[2]);
        const year = parseInt(monthMatch[3]);
        if (month && year >= 2020) {
          const result = formatDate(month, day, year);
          log(`  Found report date: ${result} (from "${monthMatch[0]}" near "${label}")`);
          return result;
        }
      }
    }

    // Pass 2: Look for MM/DD/YY or MM/DD/YYYY near labels
    for (const label of labels) {
      const labelIdx = text.toLowerCase().indexOf(label);
      if (labelIdx === -1) continue;
      const snippet = text.substring(labelIdx, labelIdx + 200);
      const dateMatch = snippet.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
      if (dateMatch) {
        const month = parseInt(dateMatch[1]);
        const day = parseInt(dateMatch[2]);
        const year = normalizeYear(dateMatch[3]);
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && year >= 2020) {
          const result = formatDate(month, day, year);
          log(`  Found report date: ${result} (near "${label}")`);
          return result;
        }
      }
    }

    // Pass 3: Find any "Month DD, YYYY" in the document
    const monthRegex = /(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+(\d{1,2}),?\s+(\d{4})/gi;
    const monthMatches = [...text.matchAll(monthRegex)];
    for (const match of monthMatches) {
      const month = monthNames[match[1].toLowerCase().substring(0, 3)];
      const day = parseInt(match[2]);
      const year = parseInt(match[3]);
      if (month && year >= 2020) {
        const result = formatDate(month, day, year);
        log(`  Found date (no label match): ${result} (from "${match[0]}")`);
        return result;
      }
    }

    // Pass 4: Find any MM/DD/YY or MM/DD/YYYY (take the last one — usually the completed date)
    const allDates = [...text.matchAll(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/g)];
    const validDates = allDates.filter(m => {
      const year = normalizeYear(m[3]);
      const month = parseInt(m[1]);
      return month >= 1 && month <= 12 && year >= 2020;
    });

    if (validDates.length > 0) {
      // Take the last date (often the "Completed" date in Labstat reports)
      const last = validDates[validDates.length - 1];
      const month = parseInt(last[1]);
      const day = parseInt(last[2]);
      const year = normalizeYear(last[3]);
      const result = formatDate(month, day, year);
      log(`  Found date (last in document): ${result}`);
      return result;
    }

    log('  No report date found in PDF.');
    return null;
  } catch (err) {
    logError('  PDF parse failed', err);
    return null;
  }
}

function calculateExpiry(reportDateStr) {
  if (!reportDateStr) return null;
  const [month, day, year] = reportDateStr.split('/').map(Number);
  const expiry = new Date(year + 1, month - 1, day);
  return `${String(expiry.getMonth() + 1).padStart(2, '0')}/${String(expiry.getDate()).padStart(2, '0')}/${expiry.getFullYear()}`;
}

// ── Strain Name Extraction ──────────────────────────────────────────────────

function extractStrainName(filename) {
  let name = filename
    .replace(/\.(pdf|png|jpe?g)$/i, '')    // Remove extension
    .replace(/_/g, ' ')                    // Replace underscores with spaces
    .replace(/\s*COA\s*(new)?\s*$/i, '')   // Remove trailing "COA" or "COA new"
    .trim();
  return name;
}

// ── Shopify File Upload ─────────────────────────────────────────────────────

async function uploadToShopify(fileBuffer, filename, mimeType = 'application/pdf') {
  const shopifyFilename = filename.replace(/\s+/g, '_');
  const store = config.shopify.store;
  const token = config.shopify.accessToken;

  // Step 1: Create staged upload
  const stagedUploadQuery = `
    mutation stagedUploadsCreate($input: [StagedUploadInput!]!) {
      stagedUploadsCreate(input: $input) {
        stagedTargets {
          url
          resourceUrl
          parameters {
            name
            value
          }
        }
        userErrors {
          field
          message
        }
      }
    }
  `;

  const stagedRes = await fetch(`https://${store}/admin/api/2024-01/graphql.json`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'X-Shopify-Access-Token': token,
    },
    body: JSON.stringify({
      query: stagedUploadQuery,
      variables: {
        input: [{
          filename: shopifyFilename,
          mimeType,
          httpMethod: 'POST',
          resource: 'FILE',
        }],
      },
    }),
  });

  const stagedData = await stagedRes.json();
  const target = stagedData.data?.stagedUploadsCreate?.stagedTargets?.[0];

  if (!target) {
    throw new Error('Failed to create staged upload: ' + JSON.stringify(stagedData));
  }

  // Step 2: Upload to staged URL
  const formData = new FormData();
  for (const param of target.parameters) {
    formData.append(param.name, param.value);
  }
  formData.append('file', new Blob([fileBuffer], { type: mimeType }), shopifyFilename);

  const uploadRes = await fetch(target.url, {
    method: 'POST',
    body: formData,
  });

  if (!uploadRes.ok) {
    throw new Error(`Upload failed: ${uploadRes.status} ${await uploadRes.text()}`);
  }

  // Step 3: Create file in Shopify
  const createFileQuery = `
    mutation fileCreate($files: [FileCreateInput!]!) {
      fileCreate(files: $files) {
        files {
          id
          alt
          ... on GenericFile {
            url
          }
        }
        userErrors {
          field
          message
        }
      }
    }
  `;

  const createRes = await fetch(`https://${store}/admin/api/2024-01/graphql.json`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'X-Shopify-Access-Token': token,
    },
    body: JSON.stringify({
      query: createFileQuery,
      variables: {
        files: [{
          originalSource: target.resourceUrl,
          contentType: 'FILE',
        }],
      },
    }),
  });

  const createData = await createRes.json();
  const fileResult = createData.data?.fileCreate?.files?.[0];

  if (!fileResult) {
    throw new Error('Failed to create file: ' + JSON.stringify(createData));
  }

  // Step 4: Poll for the file URL (it takes a moment to process)
  const fileId = fileResult.id;
  let fileUrl = null;

  for (let i = 0; i < 15; i++) {
    await new Promise(r => setTimeout(r, 2000));

    const pollQuery = `
      query {
        node(id: "${fileId}") {
          ... on GenericFile {
            url
            fileStatus
          }
        }
      }
    `;

    const pollRes = await fetch(`https://${store}/admin/api/2024-01/graphql.json`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-Shopify-Access-Token': token,
      },
      body: JSON.stringify({ query: pollQuery }),
    });

    const pollData = await pollRes.json();
    const node = pollData.data?.node;

    if (node?.url && node?.fileStatus === 'READY') {
      fileUrl = node.url;
      break;
    }

    log(`  Waiting for file processing... (attempt ${i + 1}/15)`);
  }

  if (!fileUrl) {
    throw new Error('File upload timed out waiting for processing.');
  }

  log(`  Shopify file URL: ${fileUrl}`);
  return fileUrl;
}

// ── Theme Template Update ───────────────────────────────────────────────────

function generateBlockId() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let id = 'certificate_';
  for (let i = 0; i < 6; i++) {
    id += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return id;
}

function findTargetSection(strainName, template, productType) {
  const typeConfig = config.sectionConfig[productType] || config.sectionConfig.flower;
  const sections = typeConfig.sections;

  // For non-flower types, there's typically one section
  if (sections.length === 1) {
    return sections[0].key;
  }

  // For flower types, find the right section based on alphabetical range
  const firstChar = strainName[0].toUpperCase();

  for (const sec of sections) {
    if (sec.rangeStart && sec.rangeEnd) {
      if (firstChar >= sec.rangeStart && firstChar < sec.rangeEnd) {
        return sec.key;
      }
    }
    // Handle last section (S-Z)
    if (sec === sections[sections.length - 1] && firstChar >= (sec.rangeStart || 'A')) {
      return sec.key;
    }
  }

  // Default to first section
  return sections[0].key;
}

function addCOABlock(template, strainName, fileUrl, productType = 'flower') {
  const targetSectionKey = findTargetSection(strainName, template, productType);
  const sections = template.sections;
  const typeConfig = config.sectionConfig[productType] || config.sectionConfig.flower;
  const typeSections = typeConfig.sections;
  const maxBlocks = typeConfig.maxBlocksPerSection;

  // Ensure target section exists in template
  if (!sections[targetSectionKey]) {
    throw new Error(`Target section key "${targetSectionKey}" not found in template. Check sectionConfig in config.json.`);
  }

  // Check if target section is full
  let section = sections[targetSectionKey];
  if (!section.blocks) section.blocks = {};
  if (!section.block_order) section.block_order = [];

  if (section.block_order.length >= maxBlocks) {
    log(`  ⚠️ Section "${section.name || targetSectionKey}" is full (${maxBlocks} blocks). Making room...`);

    // Find this section's index in typeSections
    const secIdx = typeSections.findIndex(s => s.key === targetSectionKey);

    // Move last block to next section (cascade)
    for (let i = secIdx; i < typeSections.length - 1; i++) {
      const currentSection = sections[typeSections[i].key];
      const nextSection = sections[typeSections[i + 1].key];

      if (!currentSection || !nextSection) {
        throw new Error(`Section key missing in template during overflow handling: ${typeSections[i].key} -> ${typeSections[i + 1].key}`);
      }

      if (!currentSection.blocks) currentSection.blocks = {};
      if (!currentSection.block_order) currentSection.block_order = [];

      if (currentSection.block_order.length >= maxBlocks) {
        // Move last block from current to beginning of next
        const lastBlockId = currentSection.block_order[currentSection.block_order.length - 1];
        const lastBlock = currentSection.blocks[lastBlockId];

        // Remove from current
        currentSection.block_order.pop();
        delete currentSection.blocks[lastBlockId];

        // Initialize blocks object if needed
        if (!nextSection.blocks) {
          nextSection.blocks = {};
        }
        if (!nextSection.block_order) {
          nextSection.block_order = [];
        }

        // Add to beginning of next (it's alphabetically the first of that section)
        nextSection.blocks[lastBlockId] = lastBlock;
        nextSection.block_order.unshift(lastBlockId);

        log(`  Moved "${lastBlock.name}" from "${currentSection.name || typeSections[i].key}" to "${nextSection.name || typeSections[i + 1].key}"`);
      }
    }

    // Check if last section overflowed
    const lastSection = sections[typeSections[typeSections.length - 1].key];
    if (!lastSection) {
      throw new Error(`Last section key "${typeSections[typeSections.length - 1].key}" not found in template.`);
    }
    if (!lastSection.block_order) lastSection.block_order = [];
    if (lastSection.block_order.length > maxBlocks) {
      log(`  ⚠️ WARNING: Last ${productType} section exceeds ${maxBlocks} blocks! Manual intervention needed.`);
    }
  }

  // Re-read the section (it may have changed)
  section = sections[targetSectionKey];

  // Create new block
  const blockId = generateBlockId();
  const newBlock = {
    type: 'certificate',
    name: strainName,
    settings: {
      certificate_name: strainName,
      certificate_file: fileUrl,
    },
  };

  // Insert alphabetically
  section.blocks[blockId] = newBlock;

  // Find correct position in block_order
  let insertIdx = section.block_order.length;
  for (let i = 0; i < section.block_order.length; i++) {
    const existingBlock = section.blocks[section.block_order[i]];
    const existingName = existingBlock?.name || existingBlock?.settings?.certificate_name || '';
    if (strainName.toLowerCase() < existingName.toLowerCase()) {
      insertIdx = i;
      break;
    }
  }

  section.block_order.splice(insertIdx, 0, blockId);
  log(`  Added "${strainName}" to section "${section.name || targetSectionKey}" at position ${insertIdx}`);

  return template;
}

async function getThemeTemplate() {
  const store = config.shopify.store;
  const token = config.shopify.accessToken;
  const themeId = config.shopify.themeId;
  const key = config.shopify.templateKey;

  const res = await fetch(
    `https://${store}/admin/api/2024-01/themes/${themeId}/assets.json?asset%5Bkey%5D=${encodeURIComponent(key)}`,
    {
      headers: {
        'X-Shopify-Access-Token': token,
        'Content-Type': 'application/json',
      },
    }
  );

  const data = await res.json();
  return JSON.parse(data.asset.value);
}

async function saveThemeTemplate(template) {
  const store = config.shopify.store;
  const token = config.shopify.accessToken;
  const themeId = config.shopify.themeId;
  const key = config.shopify.templateKey;

  const res = await fetch(
    `https://${store}/admin/api/2024-01/themes/${themeId}/assets.json`,
    {
      method: 'PUT',
      headers: {
        'X-Shopify-Access-Token': token,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        asset: {
          key,
          value: JSON.stringify(template, null, 2),
        },
      }),
    }
  );

  if (!res.ok) {
    throw new Error(`Failed to save template: ${res.status} ${await res.text()}`);
  }

  log('  Theme template updated.');
}

// ── Google Sheets ───────────────────────────────────────────────────────────

async function logToSheet(sheets, strainName, reportDate, expiryDate, shopifyUrl, driveLink, driveUrl, productType = 'flower') {
  // Map product type to category
  const categoryMap = {
    flower: 'Flower',
    edibles: 'Edibles',
    concentrates: 'Concentrates',
    prerolls: 'Pre-Rolls',
  };
  const category = categoryMap[productType] || '';
  const sheetId = config.google.sheetId;
  const sheetName = 'Sheet1';

  // Get current sheet data to find table boundaries
  const getRes = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: `'${sheetName}'!A:I`,
  });

  const values = getRes.data.values || [];

  // If no data exists, create header row and append data
  if (values.length === 0) {
    const headerRow = ['Name', 'Flower Type', 'Category', 'Report Date', 'Expiration Date', 'Posted on Retail', 'Google Drive COA Link', 'Google Drive COA URL', 'Comments'];
    await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId,
      range: `'${sheetName}'!A1:I1`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [headerRow] },
    });

    // Now append the data row with hyperlinks
    const row = [
      strainName,
      '',
      category,  // Category (smart chip)
      reportDate || '',
      expiryDate || '',
      'Yes',
      driveLink ? `=HYPERLINK("${driveLink}";"${strainName}")` : '',
      driveUrl ? `=HYPERLINK("${driveUrl}";"${driveUrl}")` : '',
      'Auto-posted by COA Automation',
    ];

    await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId,
      range: `'${sheetName}'!A2:I2`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [row] },
    });

    log('  Logged to Google Sheet (new table created).');
    return;
  }

  // Find table end (first empty row after header)
  let tableEnd = values.length;
  for (let i = 1; i < values.length; i++) {
    if (!values[i] || values[i].length === 0 || !values[i][0]) {
      tableEnd = i;
      break;
    }
  }

  // Find alphabetical insertion point
  let insertIdx = tableEnd; // Default: append at end
  const strainLower = strainName.toLowerCase();

  for (let i = 1; i < tableEnd; i++) {
    const existingName = values[i]?.[0]?.toLowerCase() || '';
    if (strainLower < existingName) {
      insertIdx = i;
      break;
    }
  }

  // Build row data with hyperlinks for Drive columns
  // Columns: Name, Flower Type, Category, Report Date, Expiration Date, Posted on Retail, Google Drive COA Link, Google Drive COA URL, Comments
  const row = [
    strainName,                         // Name (plain text)
    '',                                  // Flower Type (manual entry)
    category,                            // Category (smart chip)
    reportDate || '',                    // Report Date
    expiryDate || '',                    // Expiration Date (blank - auto-filled by formula)
    'Yes',                               // Posted on Retail
    driveLink ? `=HYPERLINK("${driveLink}";"${strainName}")` : '',  // Google Drive COA Link (smart chip - filename as link text)
    driveUrl ? `=HYPERLINK("${driveUrl}";"${driveUrl}")` : '',      // Google Drive COA URL (plain hyperlink)
    'Auto-posted by COA Automation',     // Comments
  ];

  // Get sheet ID for batchUpdate
  const sheetMetaRes = await sheets.spreadsheets.get({
    spreadsheetId: sheetId,
    fields: 'sheets.properties',
  });
  const sheet = sheetMetaRes.data.sheets.find(s => s.properties.title === sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  const googleSheetId = sheet.properties.sheetId;

  // Use batchUpdate to insert row
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: sheetId,
    requestBody: {
      requests: [
        {
          insertDimension: {
            range: {
              sheetId: googleSheetId,
              dimension: 'ROWS',
              startIndex: insertIdx,
              endIndex: insertIdx + 1,
            },
          },
        },
      ],
    },
  });

  // Populate the newly inserted row
  const dataRowNum = insertIdx + 2; // +1 for header, +1 for 1-indexing
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `'${sheetName}'!A${dataRowNum}:I${dataRowNum}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [row] },
  });

  log(`  Logged to Google Sheet (inserted at row ${dataRowNum}, alphabetically sorted).`);
}

// ── Check for Duplicate ─────────────────────────────────────────────────────

function findCOABlock(template, strainName) {
  const target = strainName.toLowerCase();
  for (const sectionKey of Object.keys(template.sections)) {
    const section = template.sections[sectionKey];
    if (!section.blocks) continue;
    for (const blockId of Object.keys(section.blocks)) {
      const block = section.blocks[blockId];
      if (block.settings?.certificate_name?.toLowerCase() === target) {
        return { sectionKey, blockId, block };
      }
    }
  }
  return null;
}

function isDuplicateInTemplate(template, strainName) {
  return Boolean(findCOABlock(template, strainName));
}

async function findRowInSheetByStrain(sheets, sheetId, sheetName, strainName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: `'${sheetName}'!A:A`,
  });
  const values = res.data.values || [];
  const target = strainName.toLowerCase();
  // values[0] is header row (A1)
  for (let i = 1; i < values.length; i++) {
    const raw = (values[i]?.[0] || '').toString().trim();
    const cell = raw.toLowerCase();

    // Column A might be plain text OR a HYPERLINK formula like:
    // =HYPERLINK("https://...";"Peaches and Cream")
    let display = cell;
    const m = raw.match(/=HYPERLINK\([^;]+;\s*"([^"]+)"\)/i);
    if (m?.[1]) display = m[1].toString().trim().toLowerCase();

    if (display === target) return i + 1; // convert 0-index to 1-index row number
  }
  return null;
}

async function updateSheetRow(sheets, sheetId, sheetName, rowNum, rowValues /* A..I */) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `'${sheetName}'!A${rowNum}:I${rowNum}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [rowValues] },
  });
  log(`  Updated Google Sheet row ${rowNum}.`);
}

// ── Main Processing ─────────────────────────────────────────────────────────

async function processFile(drive, sheets, file, processed) {
  const strainName = extractStrainName(file.name);
  const productType = file.productType || 'flower';
  log(`\nProcessing: "${file.name}" → Strain: "${strainName}" (${productType})`);

  try {
    // Load theme template once per file
    const template = await getThemeTemplate();
    const existingBlock = findCOABlock(template, strainName);

    // If the strain already exists, treat this as an UPDATE (replace the file URL)
    // instead of skipping.
    const isUpdate = Boolean(existingBlock);
    if (isUpdate) {
      log(`  ↻ "${strainName}" already exists in theme template. Will update existing entry.`);
    }

    // Download file from Drive
    log('  Downloading from Drive...');
    const fileBuffer = await downloadFile(drive, file.id);

    // Extract report date (PDF only)
    let reportDate = null;
    let expiryDate = null;
    if ((file.mimeType || '').toLowerCase() === 'application/pdf') {
      log('  Extracting report date from PDF...');
      reportDate = await extractReportDate(fileBuffer);
      expiryDate = calculateExpiry(reportDate);
      if (reportDate) {
        log(`  Report Date: ${reportDate} → Expiry: ${expiryDate}`);
      }
    } else {
      log(`  Non-PDF COA detected (${file.mimeType || 'unknown mimeType'}). Skipping PDF date extraction.`);
    }

    // Upload to Shopify
    log('  Uploading to Shopify...');
    const shopifyUrl = await uploadToShopify(fileBuffer, file.name, file.mimeType || 'application/pdf');

    // Update theme template
    log('  Updating theme template...');
    if (isUpdate && existingBlock) {
      // Replace the file URL on the existing block
      template.sections[existingBlock.sectionKey].blocks[existingBlock.blockId].settings.certificate_file = shopifyUrl;
      await saveThemeTemplate(template);
      log(`  Updated existing theme block for "${strainName}".`);
    } else {
      const updatedTemplate = addCOABlock(template, strainName, shopifyUrl, productType);
      await saveThemeTemplate(updatedTemplate);
    }

    // Log to Google Sheet (update row if exists, else insert)
    log('  Logging to Google Sheet...');
    const sheetId = config.google.sheetId;
    const sheetName = 'Sheet1';
    const existingRowNum = await findRowInSheetByStrain(sheets, sheetId, sheetName, strainName);

    const categoryMap = { flower: 'Flower', edibles: 'Edibles', concentrates: 'Concentrates', prerolls: 'Pre-Rolls' };
    const category = categoryMap[productType] || '';

    const row = [
      strainName,
      '',
      category,
      reportDate || '',
      expiryDate || '',
      'Yes',
      (file.webViewLink || '') ? `=HYPERLINK("${file.webViewLink}";"${strainName}")` : '',
      `=HYPERLINK("https://drive.google.com/file/d/${file.id}/view";"https://drive.google.com/file/d/${file.id}/view")`,
      isUpdate ? `Updated by COA Automation (${new Date().toISOString()})` : 'Auto-posted by COA Automation',
    ];

    if (existingRowNum) {
      await updateSheetRow(sheets, sheetId, sheetName, existingRowNum, row);
    } else {
      await logToSheet(
        sheets,
        strainName,
        reportDate,
        expiryDate,
        shopifyUrl,
        file.webViewLink || '',
        `https://drive.google.com/file/d/${file.id}/view`,
        productType
      );
    }

    // Mark as processed
    processed[file.id] = {
      name: file.name,
      strain: strainName,
      productType,
      reportDate,
      expiryDate,
      shopifyUrl,
      processedAt: new Date().toISOString(),
    };
    saveProcessedFiles(processed);

    log(`  ✅ Done: "${strainName}" (${productType})`);
  } catch (err) {
    logError(`Failed to process "${file.name}"`, err);
    // Don't mark as processed so it retries next time
  }
}

async function runOnce() {
  log('=== COA Automation Run ===');

  const auth = getGoogleAuth();
  const drive = google.drive({ version: 'v3', auth });
  const sheets = google.sheets({ version: 'v4', auth });
  const processed = loadProcessedFiles();

  // List new files
  const newFiles = await listNewFiles(drive, processed);
  log(`Found ${newFiles.length} new file(s) to process.`);

  if (newFiles.length === 0) {
    log('Nothing to do.');
    return;
  }

  // Process each file one at a time (to avoid race conditions on template updates)
  for (const file of newFiles) {
    await processFile(drive, sheets, file, processed);
  }

  log('\n=== Run Complete ===\n');
}

// ── Entry Point ─────────────────────────────────────────────────────────────

async function main() {
  log('COA Automation starting...');
  log(`Mode: ${ONCE_MODE ? 'single run' : 'continuous polling'}`);
  log(`Poll interval: ${config.pollIntervalMs / 1000}s`);

  if (ONCE_MODE) {
    await runOnce();
  } else {
    // Run immediately, then poll
    await runOnce();
    setInterval(async () => {
      try {
        await runOnce();
      } catch (err) {
        logError('Poll cycle failed', err);
      }
    }, config.pollIntervalMs);
  }
}

main().catch(err => {
  logError('Fatal error', err);
  process.exit(1);
});
