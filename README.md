# Expenses Automation

Automate expense reporting with Zoho Expense and Fizz invoice downloads.

## Features

- **Download Fizz invoices** - Automatically download invoices from your Fizz account
- **PDF name replacement** - Replace account holder name in downloaded PDFs using OCR
- **Create Zoho expense reports** - Read expenses from Excel and create draft reports in Zoho Expense
- **Receipt attachment** - Automatically attach PDF receipts to expenses

## Setup

### 1. Install dependencies

```bash
uv sync
uv run playwright install chromium
```

### 2. Configure environment variables

Create a `.env` file with:

```env
# Zoho Expense API
CLIENT_ID=your_zoho_client_id
CLIENT_SECRET=your_zoho_client_secret
REFRESH_TOKEN=your_zoho_refresh_token
ORGANIZATION_ID=your_zoho_org_id

# Fizz account
FIZZ_EMAIL=your_fizz_email
FIZZ_PASSWORD=your_fizz_password
FIZZ_PRODUCT_ID=your_fizz_product_id

# PDF name replacement (if needed, sometimes, the account's name is not the same as the payer's name)
PDF_OLD_NAME=Account Name
PDF_NEW_NAME=Payer Name

# Zoho IDs (from your Zoho Expense account)
ZOHO_CURRENCY_CAD=your_currency_id
ZOHO_TAX_TPS_TVQ=your_tax_id
ZOHO_CATEGORY_CONNECTIVITE=your_category_id
ZOHO_CATEGORY_ACTIVITE_PHYSIQUE=your_category_id
ZOHO_MERCHANT_FIZZ=your_merchant_id
ZOHO_MERCHANT_NAUTILUS=your_merchant_id
```

#### Getting Zoho credentials

1. Go to [Zoho API Console](https://api-console.zoho.com/)
2. Create a Self Client
3. Generate an authorization code with scopes:
   - `ZohoExpense.expensereport.READ`
   - `ZohoExpense.expensereport.CREATE`
   - `ZohoExpense.expense.CREATE`
4. Exchange the code for a refresh token:
   ```bash
   curl -X POST "https://accounts.zoho.com/oauth/v2/token" \
     -d "grant_type=authorization_code" \
     -d "client_id=YOUR_CLIENT_ID" \
     -d "client_secret=YOUR_CLIENT_SECRET" \
     -d "code=YOUR_AUTH_CODE"
   ```

#### Getting Zoho IDs

The Zoho IDs (categories, merchants, tax, currency) can be found by inspecting an existing expense report via the Zoho API or from the URL when viewing items in the Zoho Expense web interface.

#### Getting Fizz Product ID

Find your `FIZZ_PRODUCT_ID` in the URL when viewing your payment history:
```
https://zone.fizz.ca/.../wallet/payment-history?productId=YOUR_PRODUCT_ID&...
```

### 3. Set up expense tracking

Create `reports_assets/Expenses.xlsx` with tabs for each expense type:
- **Fizz** - Mobile plan data portion
- **Affirm Fizz** - Mobile plan phone portion
- **Nautilus** - Gym membership

Each tab should have columns: `Date | Submitted? | Amount` (amount in header row)

## Usage

### List existing expense reports

```bash
uv run main.py list-reports
```

### Create expense report for a month

```bash
uv run main.py create-report 2025-12
```

This will:
1. Read unsubmitted expenses from Excel for the specified month
2. Find matching PDF receipts in `reports_assets/2025-12/`
3. Create expenses in Zoho with attached receipts
4. Create a draft expense report

### Download Fizz invoice

```bash
uv run main.py download-fizz 2025-12
```

Downloads the Fizz invoice for the specified month and saves it to `reports_assets/2025-12/facture-fizz.pdf`. Automatically replaces the account holder name.

Use `--no-headless` to see the browser (for debugging):

```bash
uv run main.py download-fizz 2025-12 --no-headless
```

### Replace name in PDF (standalone)

```bash
uv run main.py replace-pdf-name reports_assets/2025-12/facture-fizz.pdf
```

Replaces `PDF_OLD_NAME` with `PDF_NEW_NAME` (configured in `.env`) in the PDF using OCR.

## Folder Structure

```
expenses/
├── main.py                 # Main script
├── .env                    # Environment variables (not committed)
├── pyproject.toml          # Project dependencies
└── reports_assets/
    ├── Expenses.xlsx       # Expense tracking spreadsheet
    ├── 2025-12/            # Month folder
    │   ├── facture-fizz.pdf
    │   └── affirm-all.pdf
    └── 2026-01/
        └── ...
```

## Receipt Matching

PDFs in month folders are matched to expense types by filename pattern:
- `*fizz*` (excluding "affirm") → Fizz expense
- `*affirm*` → Affirm Fizz expense
- Nautilus expenses don't require receipts

## Dependencies

- `playwright` - Browser automation for Fizz downloads
- `pymupdf` - PDF manipulation
- `pytesseract` - OCR for image-based PDFs
- `pandas` / `openpyxl` - Excel reading
- `requests` - Zoho API calls
- `python-dotenv` - Environment variable loading
