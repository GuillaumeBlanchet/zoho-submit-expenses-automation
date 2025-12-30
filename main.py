import asyncio
import os
import re
import sys
from pathlib import Path

import fitz  # PyMuPDF
import pandas as pd
import requests
from dotenv import load_dotenv
from playwright.async_api import async_playwright

load_dotenv()

# Zoho API configuration
ZOHO_API_DOMAIN = os.getenv("ZOHO_API_DOMAIN", "zohoapis.com")
ZOHO_ACCOUNTS_DOMAIN = os.getenv("ZOHO_ACCOUNTS_DOMAIN", "accounts.zoho.com")

# PDF name replacement
PDF_OLD_NAME = os.getenv("PDF_OLD_NAME", "")
PDF_NEW_NAME = os.getenv("PDF_NEW_NAME", "")

# Zoho IDs (from environment or defaults)
ZOHO_IDS = {
    "currency_cad": os.getenv("ZOHO_CURRENCY_CAD", ""),
    "tax_tps_tvq": os.getenv("ZOHO_TAX_TPS_TVQ", ""),
    "category_connectivite": os.getenv("ZOHO_CATEGORY_CONNECTIVITE", ""),
    "category_activite_physique": os.getenv("ZOHO_CATEGORY_ACTIVITE_PHYSIQUE", ""),
    "merchant_fizz": os.getenv("ZOHO_MERCHANT_FIZZ", ""),
    "merchant_nautilus": os.getenv("ZOHO_MERCHANT_NAUTILUS", ""),
}

# Expense type configurations mapped to Excel tab names
EXPENSE_TYPES = {
    "Fizz": {
        "category_id": ZOHO_IDS["category_connectivite"],
        "merchant_id": ZOHO_IDS["merchant_fizz"],
        "merchant_name": "Fizz",
        "description": "Partie donnée de l'abonnement.",
        "receipt_pattern": r"^(?!.*affirm).*fizz.*\.pdf$",
    },
    "Affirm Fizz": {
        "category_id": ZOHO_IDS["category_connectivite"],
        "merchant_id": ZOHO_IDS["merchant_fizz"],
        "merchant_name": "Fizz",
        "description": "Partie téléphone de l'abonnement",
        "receipt_pattern": r"affirm.*\.pdf$",
    },
    "Nautilus": {
        "category_id": ZOHO_IDS["category_activite_physique"],
        "merchant_id": ZOHO_IDS["merchant_nautilus"],
        "merchant_name": "Nautilus Plus",
        "description": "",
        "receipt_pattern": None,
    },
}

ASSETS_DIR = Path(__file__).parent / "reports_assets"


class ZohoExpenseClient:
    """Client for Zoho Expense API."""

    def __init__(self):
        self.client_id = os.getenv("CLIENT_ID")
        self.client_secret = os.getenv("CLIENT_SECRET")
        self.refresh_token = os.getenv("REFRESH_TOKEN")
        self.organization_id = os.getenv("ORGANIZATION_ID")
        self.access_token = None
        self._validate_config()

    def _validate_config(self):
        missing = []
        if not self.client_id:
            missing.append("CLIENT_ID")
        if not self.client_secret:
            missing.append("CLIENT_SECRET")
        if not self.refresh_token:
            missing.append("REFRESH_TOKEN")
        if not self.organization_id:
            missing.append("ORGANIZATION_ID")
        if missing:
            raise ValueError(f"Missing required env vars: {', '.join(missing)}")

    def authenticate(self):
        """Get a fresh access token."""
        url = f"https://{ZOHO_ACCOUNTS_DOMAIN}/oauth/v2/token"
        params = {
            "grant_type": "refresh_token",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "refresh_token": self.refresh_token,
        }
        response = requests.post(url, params=params)
        response.raise_for_status()
        data = response.json()
        if "error" in data:
            raise Exception(f"OAuth error: {data['error']}")
        self.access_token = data["access_token"]
        return self.access_token

    def _headers(self):
        return {
            "Authorization": f"Zoho-oauthtoken {self.access_token}",
            "X-com-zoho-expense-organizationid": self.organization_id,
        }

    def list_reports(self) -> list:
        """List all expense reports."""
        url = f"https://www.{ZOHO_API_DOMAIN}/expense/v1/expensereports"
        response = requests.get(url, headers=self._headers())
        response.raise_for_status()
        return response.json().get("expense_reports", [])

    def create_expense(
        self,
        date: str,
        amount: float,
        category_id: str,
        merchant_id: str,
        description: str = "",
        receipt_path: Path | None = None,
    ) -> dict:
        """Create an expense, optionally with a receipt attachment."""
        url = f"https://www.{ZOHO_API_DOMAIN}/expense/v1/expenses"

        expense_data = {
            "currency_id": ZOHO_IDS["currency_cad"],
            "date": date,
            "is_reimbursable": True,
            "is_inclusive_tax": True,
            "line_items": [
                {
                    "category_id": category_id,
                    "amount": amount,
                    "tax_id": ZOHO_IDS["tax_tps_tvq"],
                    "description": description,
                }
            ],
            "merchant_id": merchant_id,
        }

        if receipt_path and receipt_path.exists():
            # Multipart upload with receipt
            import json

            files = {
                "receipt": (receipt_path.name, open(receipt_path, "rb"), "application/pdf"),
            }
            data = {"JSONString": json.dumps(expense_data)}
            response = requests.post(
                url,
                headers=self._headers(),
                data=data,
                files=files,
            )
        else:
            # JSON-only request
            response = requests.post(
                url,
                headers=self._headers(),
                json=expense_data,
            )

        response.raise_for_status()
        result = response.json()
        if result.get("code") != 0:
            raise Exception(f"API error: {result.get('message')}")
        # API returns "expenses" array, not singular "expense"
        expenses = result.get("expenses", [])
        return expenses[0] if expenses else {}

    def create_report(
        self,
        report_name: str,
        start_date: str,
        end_date: str,
        expense_ids: list[str],
    ) -> dict:
        """Create a draft expense report with the given expenses."""
        url = f"https://www.{ZOHO_API_DOMAIN}/expense/v1/expensereports"

        report_data = {
            "report_name": report_name,
            "start_date": start_date,
            "end_date": end_date,
            "expenses": [
                {"expense_id": eid, "order": i} for i, eid in enumerate(expense_ids)
            ],
        }

        response = requests.post(url, headers=self._headers(), json=report_data)
        response.raise_for_status()
        result = response.json()
        if result.get("code") != 0:
            raise Exception(f"API error: {result.get('message')}")
        return result.get("expense_report", {})


class FizzInvoiceDownloader:
    """Download Fizz invoices using browser automation."""

    FIZZ_LOGIN_URL = "https://zone.fizz.ca"

    def __init__(self):
        self.email = os.getenv("FIZZ_EMAIL")
        self.password = os.getenv("FIZZ_PASSWORD")
        self.product_id = os.getenv("FIZZ_PRODUCT_ID")
        self._validate_config()

    FRENCH_MONTHS = {
        "01": ["janv", "janvier"],
        "02": ["févr", "février"],
        "03": ["mars"],
        "04": ["avr", "avril"],
        "05": ["mai"],
        "06": ["juin"],
        "07": ["juil", "juillet"],
        "08": ["août"],
        "09": ["sept", "septembre"],
        "10": ["oct", "octobre"],
        "11": ["nov", "novembre"],
        "12": ["déc", "décembre"],
    }

    @property
    def payment_history_url(self):
        return f"https://zone.fizz.ca/dce/customer-ui-prod/wallet/payment-history?productId={self.product_id}&prvState=plan"

    def _validate_config(self):
        missing = []
        if not self.email:
            missing.append("FIZZ_EMAIL")
        if not self.password:
            missing.append("FIZZ_PASSWORD")
        if missing:
            raise ValueError(f"Missing required env vars: {', '.join(missing)}")

    async def download_invoice(self, month: str, output_path: Path, headless: bool = True) -> bool:
        """
        Download Fizz invoice for the given month.
        Returns True on success, False on failure.
        """
        year, month_num = month.split("-")

        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=headless)
            context = await browser.new_context(
                user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            )
            page = await context.new_page()

            try:
                # Step 1: Navigate to payment history (will redirect to login if needed)
                print("  Navigating to Fizz...")
                await page.goto(self.payment_history_url, wait_until="domcontentloaded")
                await page.wait_for_timeout(5000)

                # Check if we're on login page
                current_url = page.url
                print(f"  Current URL: {current_url}")

                # Handle Okta OAuth login if redirected to auth.fizz.ca
                if "auth.fizz.ca" in current_url or "okta" in current_url.lower():
                    print("  Detected Okta login page, waiting for form...")

                    # Wait for username field to be visible
                    try:
                        await page.wait_for_selector('input[name="identifier"], input[name="username"], input[id="okta-signin-username"]', timeout=10000)
                    except Exception:
                        # Take screenshot to debug
                        debug_path = ASSETS_DIR / "debug_login.png"
                        await page.screenshot(path=str(debug_path))
                        print(f"  Saved debug screenshot to {debug_path}")
                        raise

                    # Fill username/email
                    username_input = await page.query_selector('input[name="identifier"], input[name="username"], input[id="okta-signin-username"]')
                    if username_input:
                        print("  Filling email...")
                        await username_input.fill(self.email)

                    # Check if password field is already visible (single page login)
                    password_input = await page.query_selector('input[type="password"]:visible')
                    if password_input:
                        print("  Filling password (single page)...")
                        await password_input.fill(self.password)
                    else:
                        # Click Next for two-step login
                        next_btn = await page.query_selector('input[type="submit"], button[type="submit"]')
                        if next_btn:
                            print("  Clicking next...")
                            await next_btn.click()
                            await page.wait_for_timeout(3000)

                        # Wait for and fill password
                        try:
                            await page.wait_for_selector('input[type="password"]', timeout=10000)
                            password_input = await page.query_selector('input[type="password"]')
                            if password_input:
                                print("  Filling password...")
                                await password_input.fill(self.password)
                        except Exception:
                            debug_path = ASSETS_DIR / "debug_password.png"
                            await page.screenshot(path=str(debug_path))
                            print(f"  Saved debug screenshot to {debug_path}")
                            raise

                    # Click sign in
                    submit_btn = await page.query_selector('input[type="submit"], button[type="submit"]')
                    if submit_btn:
                        print("  Submitting login...")
                        await submit_btn.click()

                    # Wait for redirect after login (with longer timeout)
                    print("  Waiting for login to complete...")
                    try:
                        await page.wait_for_url("**/zone.fizz.ca/**", timeout=30000)
                    except Exception:
                        # May have landed on different page, continue anyway
                        pass

                    await page.wait_for_timeout(3000)
                    print(f"  After login URL: {page.url}")

                # If still not on payment history, navigate there
                if "payment-history" not in page.url:
                    print("  Navigating to payment history...")
                    await page.goto(self.payment_history_url, wait_until="domcontentloaded")
                    await page.wait_for_load_state("networkidle")
                    await page.wait_for_timeout(3000)

                print(f"  On page: {page.url}")

                # Wait for page content to load
                await page.wait_for_timeout(3000)

                # Step 2: Find and click the row for the target month
                print(f"  Looking for invoice for {month}...")

                # Build patterns to match date formats: YYYY/MM, YYYY-MM, or French month names
                french_months = self.FRENCH_MONTHS.get(month_num, [])
                numeric_patterns = [
                    f"{year}/{month_num}",  # 2025/12
                    f"{year}-{month_num}",  # 2025-12
                ]

                def matches_month(text: str) -> bool:
                    """Check if text contains the target month."""
                    text_lower = text.lower()
                    # Check numeric patterns
                    for pattern in numeric_patterns:
                        if pattern in text:
                            return True
                    # Check French month names
                    if year in text and any(m in text_lower for m in french_months):
                        return True
                    return False

                # Look for clickable elements containing the month
                # Try table rows first (most likely structure)
                selectors_to_try = [
                    'table tbody tr',
                    'table tr',
                    'div[class*="transaction"]',
                    'div[class*="row"]',
                    'div[class*="item"]',
                    'li',
                ]

                target_row = None
                for selector in selectors_to_try:
                    rows = await page.query_selector_all(selector)
                    if rows:
                        for row in rows:
                            text = await row.inner_text()
                            if matches_month(text):
                                target_row = row
                                print(f"  Found matching row: {text[:100].strip()}...")
                                break
                        if target_row:
                            break

                if not target_row:
                    print(f"  ERROR: Could not find invoice for {month}")
                    # Save screenshot for debugging
                    debug_path = ASSETS_DIR / "debug_payment_history.png"
                    await page.screenshot(path=str(debug_path), full_page=True)
                    print(f"  Saved debug screenshot to {debug_path}")
                    return False

                # Step 4: Click the row to open modal
                print("  Clicking to open invoice...")
                await target_row.click()
                await page.wait_for_timeout(2000)

                # Step 5: Download the PDF
                print("  Looking for PDF download...")

                # Try multiple approaches to find/download PDF
                # Approach 1: Look for download link
                download_link = await page.query_selector('a[href*=".pdf"], a[download], button:has-text("Download"), button:has-text("Télécharger")')
                if download_link:
                    async with page.expect_download() as download_info:
                        await download_link.click()
                    download = await download_info.value
                    await download.save_as(output_path)
                    print(f"  Downloaded via link: {output_path}")
                    return True

                # Approach 2: Look for PDF in iframe
                iframe = await page.query_selector('iframe[src*=".pdf"], iframe[src*="pdf"]')
                if iframe:
                    src = await iframe.get_attribute("src")
                    if src:
                        # Download PDF directly
                        response = await page.request.get(src)
                        content = await response.body()
                        output_path.write_bytes(content)
                        print(f"  Downloaded from iframe: {output_path}")
                        return True

                # Approach 3: Look for embedded PDF object
                pdf_object = await page.query_selector('object[type="application/pdf"], embed[type="application/pdf"]')
                if pdf_object:
                    src = await pdf_object.get_attribute("data") or await pdf_object.get_attribute("src")
                    if src:
                        response = await page.request.get(src)
                        content = await response.body()
                        output_path.write_bytes(content)
                        print(f"  Downloaded from embed: {output_path}")
                        return True

                print("  ERROR: Could not find PDF download mechanism")
                return False

            except Exception as e:
                print(f"  ERROR: {e}")
                return False
            finally:
                await browser.close()


def replace_name_in_pdf(pdf_path: Path, old_name: str, new_name: str) -> bool:
    """
    Replace a name in a PDF file.
    Handles both text-based and image-based PDFs.
    Returns True on success, False on failure.
    """
    try:
        doc = fitz.open(pdf_path)
        found_any = False

        for page_num in range(len(doc)):
            page = doc[page_num]

            # First try text-based search
            text_instances = page.search_for(old_name)

            if text_instances:
                found_any = True
                print(f"  Found {len(text_instances)} instance(s) of '{old_name}' on page {page_num + 1}")

                for inst in text_instances:
                    page.add_redact_annot(inst, fill=(1, 1, 1))
                page.apply_redactions()

                for inst in text_instances:
                    font_size = inst.height * 0.85
                    page.insert_text(
                        point=(inst.x0, inst.y1 - 2),
                        text=new_name,
                        fontsize=font_size,
                        fontname="helv",
                        color=(0, 0, 0),
                    )
            else:
                # Try OCR for image-based PDFs
                found_any = _replace_name_in_image_pdf(page, old_name, new_name)

        if not found_any:
            print(f"  No instances of '{old_name}' found in PDF")
            doc.close()
            return True

        # Save to a temporary file first, then replace original
        temp_path = pdf_path.with_suffix(".tmp.pdf")
        doc.save(temp_path, deflate=True)
        doc.close()

        # Replace original with modified version
        temp_path.replace(pdf_path)
        return True

    except Exception as e:
        print(f"  ERROR: Failed to modify PDF: {e}")
        return False


def _replace_name_in_image_pdf(page, old_name: str, new_name: str) -> bool:
    """
    Use OCR to find and replace text in an image-based PDF page.
    Returns True if text was found and replaced.
    """
    import pytesseract
    from PIL import Image
    import io

    # Render page to high-resolution image
    zoom = 2  # 2x zoom for better OCR accuracy
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    img_data = pix.tobytes("png")
    img = Image.open(io.BytesIO(img_data))

    # Run OCR with bounding box info
    ocr_data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT, lang='fra+eng')

    # Find the old name in OCR results
    found = False
    words = ocr_data['text']
    n_boxes = len(words)

    # Look for consecutive words that match the old name
    old_name_parts = old_name.split()

    for i in range(n_boxes - len(old_name_parts) + 1):
        # Check if consecutive words match
        match = True
        for j, part in enumerate(old_name_parts):
            if i + j >= n_boxes or words[i + j].lower() != part.lower():
                match = False
                break

        if match:
            found = True
            print(f"  Found '{old_name}' via OCR on page")

            # Get bounding box for all matched words (in image coordinates)
            x0 = ocr_data['left'][i]
            y0 = ocr_data['top'][i]
            x1 = ocr_data['left'][i + len(old_name_parts) - 1] + ocr_data['width'][i + len(old_name_parts) - 1]
            y1 = max(ocr_data['top'][j] + ocr_data['height'][j] for j in range(i, i + len(old_name_parts)))

            # Convert from image coordinates to PDF coordinates (account for zoom)
            pdf_x0 = x0 / zoom
            pdf_y0 = y0 / zoom
            pdf_x1 = x1 / zoom
            pdf_y1 = y1 / zoom

            # Get font height from OCR - use larger size for bolder appearance
            font_height = (pdf_y1 - pdf_y0) * 1.15

            # Draw white rectangle to cover old text (with some padding)
            rect = fitz.Rect(pdf_x0 - 2, pdf_y0 - 2, pdf_x1 + 2, pdf_y1 + 2)
            page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1))

            # Insert new text with bold font
            page.insert_text(
                point=(pdf_x0, pdf_y1 - 1),
                text=new_name,
                fontsize=font_height,
                fontname="hebo",  # Helvetica Bold
                color=(0, 0, 0),
            )

    return found


def cmd_replace_pdf_name(pdf_path: str, old_name: str | None = None, new_name: str | None = None):
    """Replace a name in a PDF file."""
    path = Path(pdf_path)
    if not path.exists():
        print(f"Error: File not found: {pdf_path}")
        return

    # Use environment variables if not provided
    old_name = old_name or PDF_OLD_NAME
    new_name = new_name or PDF_NEW_NAME

    print(f"Replacing '{old_name}' with '{new_name}' in {path.name}...")
    success = replace_name_in_pdf(path, old_name, new_name)

    if success:
        print(f"Successfully updated {path}")
    else:
        print(f"Failed to update {path}")


def read_expenses_from_excel(month: str) -> list[dict]:
    """
    Read expenses from Excel for a given month (YYYY-MM format).
    Returns list of expense dicts with: type, date, amount
    """
    excel_path = ASSETS_DIR / "Expenses.xlsx"
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    year, month_num = month.split("-")
    expenses = []

    xlsx = pd.ExcelFile(excel_path)
    for sheet_name in xlsx.sheet_names:
        if sheet_name not in EXPENSE_TYPES:
            print(f"  Warning: Unknown sheet '{sheet_name}', skipping")
            continue

        df = pd.read_excel(xlsx, sheet_name=sheet_name)

        # The Excel has amount in the header (column index 2)
        # Get amount from column name
        amount_col = df.columns[2] if len(df.columns) > 2 else None
        if amount_col is not None and isinstance(amount_col, (int, float)):
            amount = float(amount_col)
        else:
            print(f"  Warning: Could not find amount for '{sheet_name}'")
            continue

        # Filter by month and Submitted? = No
        date_col = df.columns[0]
        submitted_col = df.columns[1]

        for _, row in df.iterrows():
            expense_date = pd.to_datetime(row[date_col])
            submitted = str(row[submitted_col]).strip().lower()

            # Check if date is in target month and not submitted
            if (
                expense_date.strftime("%Y-%m") == month
                and submitted == "no"
            ):
                expenses.append({
                    "type": sheet_name,
                    "date": expense_date.strftime("%Y-%m-%d"),
                    "amount": amount,
                })

    return expenses


def find_receipt_for_expense(month_folder: Path, expense_type: str) -> Path | None:
    """Find a matching receipt PDF for the expense type."""
    config = EXPENSE_TYPES.get(expense_type)
    if not config or not config.get("receipt_pattern"):
        return None

    if not month_folder.exists():
        return None

    pattern = re.compile(config["receipt_pattern"], re.IGNORECASE)
    for pdf_file in month_folder.glob("*.pdf"):
        if pattern.search(pdf_file.name):
            return pdf_file

    return None


def cmd_download_fizz(month: str, headless: bool = True):
    """Download Fizz invoice for the given month."""
    print(f"Downloading Fizz invoice for {month}...\n")

    # Create month folder if needed
    month_folder = ASSETS_DIR / month
    month_folder.mkdir(parents=True, exist_ok=True)

    output_path = month_folder / "facture-fizz.pdf"

    downloader = FizzInvoiceDownloader()
    success = asyncio.run(downloader.download_invoice(month, output_path, headless=headless))

    if success:
        print(f"\nInvoice saved to: {output_path}")

        # Replace name in PDF
        print("\nUpdating name in PDF...")
        replace_name_in_pdf(output_path, PDF_OLD_NAME, PDF_NEW_NAME)
        print("Done!")
    else:
        print(f"\nFailed to download invoice. Try running with --no-headless to debug.")


def cmd_list_reports():
    """List existing expense reports."""
    client = ZohoExpenseClient()
    print("Authenticating with Zoho...")
    client.authenticate()
    print("✓ Successfully authenticated\n")

    print("Fetching expense reports...")
    reports = client.list_reports()

    if not reports:
        print("No expense reports found")
        return

    print(f"✓ Found {len(reports)} expense report(s):\n")
    for report in reports:
        status = report.get("status", "Unknown")
        total = report.get("total", 0)
        currency = report.get("currency_code", "")
        print(f"  [{status}] {report.get('report_number', 'Unnamed')} - {report.get('report_name', '')}")
        print(f"          Total: {currency} {total}")
        print(f"          ID: {report.get('report_id')}")
        print()


def cmd_create_report(month: str):
    """Create a draft expense report for the given month."""
    print(f"Creating expense report for {month}...\n")

    # Read expenses from Excel
    print("Reading expenses from Excel...")
    expenses = read_expenses_from_excel(month)
    if not expenses:
        print(f"No unsubmitted expenses found for {month}")
        return

    print(f"✓ Found {len(expenses)} expense(s) to create:\n")
    for exp in expenses:
        print(f"  - {exp['type']}: {exp['date']} - ${exp['amount']:.2f}")
    print()

    # Find receipts
    month_folder = ASSETS_DIR / month
    print(f"Looking for receipts in {month_folder}...")
    receipt_map = {}
    for exp in expenses:
        receipt = find_receipt_for_expense(month_folder, exp["type"])
        if receipt:
            receipt_map[f"{exp['type']}_{exp['date']}"] = receipt
            print(f"  ✓ Found receipt for {exp['type']}: {receipt.name}")
        elif EXPENSE_TYPES[exp["type"]].get("receipt_pattern"):
            print(f"  ! No receipt found for {exp['type']}")
    print()

    # Initialize client and authenticate
    client = ZohoExpenseClient()
    print("Authenticating with Zoho...")
    client.authenticate()
    print("✓ Successfully authenticated\n")

    # Create expenses
    print("Creating expenses in Zoho...")
    expense_ids = []
    for exp in expenses:
        config = EXPENSE_TYPES[exp["type"]]
        receipt_key = f"{exp['type']}_{exp['date']}"
        receipt_path = receipt_map.get(receipt_key)

        try:
            created = client.create_expense(
                date=exp["date"],
                amount=exp["amount"],
                category_id=config["category_id"],
                merchant_id=config["merchant_id"],
                description=config["description"],
                receipt_path=receipt_path,
            )
            expense_id = created.get("expense_id")
            expense_ids.append(expense_id)
            receipt_status = f" (with {receipt_path.name})" if receipt_path else ""
            print(f"  ✓ Created {exp['type']} expense: {expense_id}{receipt_status}")
        except Exception as e:
            print(f"  ✗ Failed to create {exp['type']} expense: {e}")

    if not expense_ids:
        print("\nNo expenses were created, cannot create report")
        return

    # Create report
    print("\nCreating draft expense report...")
    dates = [exp["date"] for exp in expenses]
    start_date = min(dates)
    end_date = max(dates)

    try:
        report = client.create_report(
            report_name=f"Expenses {month}",
            start_date=start_date,
            end_date=end_date,
            expense_ids=expense_ids,
        )
        report_number = report.get("report_number", "Unknown")
        report_id = report.get("report_id")
        print(f"\n✓ Created draft report: {report_number}")
        print(f"  Report ID: {report_id}")
        print(f"  Date range: {start_date} to {end_date}")
        print(f"  Expenses: {len(expense_ids)}")
        print(f"\n  View at: https://expense.zoho.com/app/{client.organization_id}#/expensereports/{report_id}")
    except Exception as e:
        print(f"\n✗ Failed to create report: {e}")
        print("  Expenses were created but not added to a report")
        print(f"  Expense IDs: {expense_ids}")


def print_usage():
    print(f"""
Usage: uv run main.py <command> [args]

Commands:
  list-reports                    List existing expense reports
  create-report <YYYY-MM>         Create a draft report for the given month
  download-fizz <YYYY-MM>         Download Fizz invoice for the given month
  download-fizz <YYYY-MM> --no-headless   Run with visible browser (for debugging)
  replace-pdf-name <file.pdf>     Replace PDF_OLD_NAME with PDF_NEW_NAME in PDF

Examples:
  uv run main.py list-reports
  uv run main.py create-report 2025-12
  uv run main.py download-fizz 2025-12
  uv run main.py replace-pdf-name reports_assets/2025-12/facture-fizz.pdf

Current PDF name replacement: "{PDF_OLD_NAME}" -> "{PDF_NEW_NAME}"
""")


def main():
    if len(sys.argv) < 2:
        print_usage()
        return

    command = sys.argv[1]

    try:
        if command == "list-reports":
            cmd_list_reports()
        elif command == "create-report":
            if len(sys.argv) < 3:
                print("Error: Missing month argument (YYYY-MM)")
                print_usage()
                return
            month = sys.argv[2]
            if not re.match(r"^\d{4}-\d{2}$", month):
                print(f"Error: Invalid month format '{month}'. Use YYYY-MM")
                return
            cmd_create_report(month)
        elif command == "download-fizz":
            if len(sys.argv) < 3:
                print("Error: Missing month argument (YYYY-MM)")
                print_usage()
                return
            month = sys.argv[2]
            if not re.match(r"^\d{4}-\d{2}$", month):
                print(f"Error: Invalid month format '{month}'. Use YYYY-MM")
                return
            headless = "--no-headless" not in sys.argv
            cmd_download_fizz(month, headless=headless)
        elif command == "replace-pdf-name":
            if len(sys.argv) < 3:
                print("Error: Missing PDF file path")
                print_usage()
                return
            pdf_path = sys.argv[2]
            cmd_replace_pdf_name(pdf_path)
        else:
            print(f"Unknown command: {command}")
            print_usage()
    except ValueError as e:
        print(f"Configuration error: {e}")
        print("\nMake sure your .env file contains:")
        print("  CLIENT_ID=...")
        print("  CLIENT_SECRET=...")
        print("  REFRESH_TOKEN=...")
        print("  ORGANIZATION_ID=...")
        print("\nFor Fizz downloads, also add:")
        print("  FIZZ_EMAIL=...")
        print("  FIZZ_PASSWORD=...")
    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()
