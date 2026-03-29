# COA Automation - Technical Notes

## Current Structure

The COA page uses a **custom Shopify template** (`page.certificate-of-analysis.json`) with JSON sections.

### Sections (in order):
1. **Flower COAs** (section key defined in `config.json`) - A-S
2. **Flower COAs 2** (section key defined in `config.json`) - S-Z
3. **Pre-Rolls COAs** (section key defined in `config.json`)
4. **Concentrate COAs** (section key defined in `config.json`)
5. **Edibles COAs** (section key defined in `config.json`)

### Block Structure:
```json
{
  "certificate_XXXXXX": {
    "type": "certificate",
    "name": "Strain Name",
    "settings": {
      "certificate_name": "Strain Name",
      "product_image": "shopify://shop_images/image.png",
      "certificate_file": "https://cdn.shopify.com/s/files/..."
    }
  }
}
```

### Key Observations:
- Block IDs follow pattern: `certificate_` + 6 random alphanumeric chars
- Blocks are alphabetically ordered within sections
- Flower COAs are split across multiple sections (due to Shopify block limits per section)
- `product_image` is optional (some entries don't have it)
- `certificate_file` links to Shopify CDN (uploaded via Files API)
- `show_images` is set to `false` in all sections

## Configuration

All IDs (theme ID, store URL, Drive folder IDs, Sheet ID) live in `config.json`.
Copy `config.example.json` → `config.json` and fill in your values.

## Google Drive

Watches one or more Drive folders (configurable per product type) for new COA PDFs or images.

## Google Sheet Columns
1. Name
2. Flower Type
3. Category
4. Report Date
5. Expiration Date
6. Posted on Retail
7. Google Drive COA Link
8. Google Drive COA URL
9. Comments

## Filename Patterns
- `"Tres Leches.pdf"` — plain strain name
- `"Kush Mintz COA.pdf"` — trailing "COA" is stripped automatically
- Shopify replaces spaces with underscores in filenames on upload
