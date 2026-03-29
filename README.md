# Certificate of Analysis Processor

An automated backend service that monitors Google Drive for new Certificate of Analysis (COA) documents, extracts metadata from them, uploads them to a Shopify storefront, updates the live theme template, and logs everything to a Google Sheet — all without human intervention.

Built to solve a real operational problem for a cannabis dispensary: manually uploading lab results was slow, error-prone, and didn't scale. This service turned a multi-step manual process into a fully automated pipeline triggered by simply dropping a file into a Drive folder.

---

## How It Works

```
Google Drive Folder
       │
       ▼
  Detect new COA file (PDF or image)
       │
       ▼
  Download file & extract report date (PDF parsing)
       │
       ▼
  Upload file to Shopify via Staged Uploads API
       │
       ▼
  Fetch live Shopify theme template (JSON)
  Insert/update COA block alphabetically
  Push updated template back to Shopify
       │
       ▼
  Log entry to Google Sheet (alphabetically sorted)
       │
       ▼
  Mark file as processed (local JSON cache)
```

The service runs as a persistent daemon (via `systemd`) and polls Drive on a configurable interval. It can also be triggered for a single run with `--once`.

---

## Features

- **Multi-format support** — handles PDF, PNG, and JPEG COA files
- **Intelligent PDF date extraction** — uses a multi-pass regex strategy to find report dates across various lab report formats (labeled dates → unlabeled dates → last date in document)
- **Automatic expiry calculation** — computes a 1-year expiry from the report date
- **Alphabetical insertion** — COA blocks are inserted in alphabetical order into both the Shopify theme template and the Google Sheet
- **Overflow/cascade handling** — if a Shopify theme section reaches its block limit, the last entry is automatically cascaded into the next section
- **Update vs. insert logic** — if a strain already exists, the existing Shopify block and sheet row are updated in place rather than duplicating
- **Multi-category support** — separate Drive folders and theme sections for flower, edibles, concentrates, and pre-rolls
- **Token auto-refresh** — Google OAuth tokens are refreshed and persisted automatically
- **Idempotent processing** — a local `processed-files.json` cache prevents re-processing the same file across runs
- **Systemd-ready** — includes a `.service` file for running as a persistent background service on Linux

---

## Tech Stack

| Layer | Technology |
|---|---|
| Runtime | Node.js (ES Modules) |
| Google APIs | `googleapis` (Drive v3, Sheets v4) |
| PDF Parsing | `pdf-parse` |
| Shopify | Admin REST API 2024-01 + GraphQL (Staged Uploads, Files, Themes) |
| Auth | Google OAuth 2.0 (installed app flow) |
| Process Management | systemd |

---

## Project Structure

```
coa-automation/
├── index.js                       # Main application logic
├── package.json
├── config.example.json            # Configuration template (copy → config.json)
├── google-credentials.example.json  # Google OAuth credentials template
├── coa-automation.service         # systemd service unit file
├── NOTES.md                       # Developer notes on template structure
└── .gitignore
```

---

## Prerequisites

- Node.js v18+
- A Google Cloud project with the Drive API and Sheets API enabled
- OAuth 2.0 credentials (Desktop app type) downloaded as `google-credentials.json`
- A Shopify store with a private app access token (`read_themes`, `write_themes`, `read_files`, `write_files` scopes)
- A Shopify theme using a custom `page.certificate-of-analysis.json` template with `certificate` block types
- Google Drive folders set up for each product category
- A Google Sheet for logging

---

## Installation

**1. Clone the repository**

```bash
git clone https://github.com/your-username/certificate-of-analysis-processor.git
cd certificate-of-analysis-processor
```

**2. Install dependencies**

```bash
npm install
```

**3. Configure credentials**

```bash
cp config.example.json config.json
cp google-credentials.example.json google-credentials.json
```

Open `config.json` and fill in your values:

```json
{
  "shopify": {
    "store": "your-store.myshopify.com",
    "accessToken": "shpat_...",
    "themeId": 000000000000,
    "templateKey": "templates/page.certificate-of-analysis.json"
  },
  "google": {
    "credentialsPath": "./google-credentials.json",
    "tokenPath": "./token.json",
    "driveMainFolderId": "...",
    "driveFolders": {
      "flower": "...",
      "edibles": "...",
      "concentrates": "..."
    },
    "sheetId": "..."
  },
  "pollIntervalMs": 300000,
  "processedFilesPath": "./processed-files.json",
  "sectionConfig": { ... }
}
```

Open `google-credentials.json` and paste your OAuth client credentials from Google Cloud Console.

**4. Authenticate with Google**

```bash
npm run setup
```

This runs the OAuth flow and saves a `token.json` locally. You only need to do this once.

**5. Run**

Single run (process all new files and exit):

```bash
npm run once
```

Continuous polling (runs every `pollIntervalMs` milliseconds):

```bash
npm start
```

---

## Running as a systemd Service (Linux/VPS)

**1. Update the service file**

Edit `coa-automation.service` and set `WorkingDirectory` to the absolute path of this project.

**2. Install and enable**

```bash
sudo cp coa-automation.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable coa-automation
sudo systemctl start coa-automation
```

**3. Check logs**

```bash
journalctl -u coa-automation -f
```

---

## Configuration Reference

### `config.json`

| Key | Description |
|---|---|
| `shopify.store` | Your Shopify store domain (e.g. `your-store.myshopify.com`) |
| `shopify.accessToken` | Private app Admin API access token |
| `shopify.themeId` | Numeric ID of the theme to update |
| `shopify.templateKey` | Path to the COA template within the theme (e.g. `templates/page.certificate-of-analysis.json`) |
| `google.credentialsPath` | Path to your `google-credentials.json` file |
| `google.tokenPath` | Path where the OAuth token will be saved |
| `google.driveMainFolderId` | Parent Drive folder ID (informational) |
| `google.driveFolders` | Object mapping product type → Drive folder ID |
| `google.sheetId` | Google Sheet ID for logging |
| `pollIntervalMs` | Polling interval in milliseconds (default: 300000 = 5 min) |
| `processedFilesPath` | Path to the local processed-files cache |
| `sectionConfig` | Maps each product type to its theme section key(s) and block limits |

### `sectionConfig` Detail

Each product type maps to one or more theme sections. For product types with many entries (like flower), multiple sections can be defined with alphabetical ranges to work around Shopify's per-section block limits.

```json
"flower": {
  "sections": [
    { "key": "section_key_1", "name": "Flower COAs A-S", "rangeStart": "A", "rangeEnd": "S" },
    { "key": "section_key_2", "name": "Flower COAs S-Z", "rangeStart": "S", "rangeEnd": "Z" }
  ],
  "maxBlocksPerSection": 200
}
```

---

## Google Sheet Schema

The sheet is auto-created with headers on first run. Columns (A–I):

| Column | Content |
|---|---|
| A | Strain / Product Name |
| B | Flower Type (manual entry) |
| C | Category (Flower / Edibles / Concentrates / Pre-Rolls) |
| D | Report Date (extracted from PDF) |
| E | Expiration Date (Report Date + 1 year) |
| F | Posted on Retail (always "Yes") |
| G | Google Drive COA Link (hyperlink) |
| H | Google Drive COA URL (plain hyperlink) |
| I | Comments (auto-generated note) |

Rows are inserted alphabetically by strain name.

---

## Date Extraction Logic

The PDF parser uses a four-pass strategy to find the report date, designed to handle the varying formats used by different cannabis testing labs:

1. **Labeled month-name dates** — looks for "Month DD, YYYY" within 200 characters of a date-related label (e.g., "Completed", "Report Date", "Date Issued")
2. **Labeled numeric dates** — looks for `MM/DD/YY` or `MM/DD/YYYY` near the same labels
3. **Any month-name date** — falls back to the first valid "Month DD, YYYY" anywhere in the document
4. **Any numeric date** — takes the last valid `MM/DD/YYYY` in the document (typically the "Completed" date in most lab formats)

---

## License

[CC BY-NC 4.0](LICENSE) — Free to use and adapt for non-commercial purposes with attribution.
