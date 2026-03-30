import streamlit as st
import re
import io
from pdf2image import convert_from_path
import pytesseract
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import tempfile
import os

st.set_page_config(page_title="Bank Register to Excel", layout="wide")
st.title("Bank Register PDF to Excel Converter")
st.write("Upload an Atlantic Union Bank register PDF to convert it into a structured Excel file with debits, credits, and running balances.")

MONTH_MAP = {
    'JAN': 1, 'FEB': 2, 'MAR': 3, 'APR': 4, 'MAY': 5, 'JUN': 6,
    'JUL': 7, 'AUG': 8, 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DEC': 12
}
MONTH_NAMES = '|'.join(MONTH_MAP.keys())
MONTH_PAT = rf'({MONTH_NAMES})'

# Known description keywords for validation
DESC_KEYWORDS = ['TUITION', 'FIDELITY', 'COMMONWEALTH', 'COVA', 'NETWORK',
                 'PAYPAL', 'BANKCARD', 'FUNDS TRANSFER', 'DEPOSIT', 'TRANSFER',
                 'PMT', 'PAYMENT', 'INVESTM', 'VENDORPAYM']


def ocr_pdf_to_images(uploaded_file):
    """Convert uploaded PDF to page images."""
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(uploaded_file.read())
        tmp_path = tmp.name
    try:
        return convert_from_path(tmp_path, dpi=300)
    finally:
        os.unlink(tmp_path)


def parse_amount(text):
    """Parse a dollar amount string. Parentheses = negative (debit)."""
    text = text.strip().rstrip(';:. ')
    # ($1,234.56) -> negative
    neg = re.match(r'^\(\$?([\d,]+\.\d{2})\)$', text)
    if neg:
        return -float(neg.group(1).replace(',', ''))
    # $1,234.56 -> positive
    pos = re.match(r'^\$?([\d,]+\.\d{2})$', text)
    if pos:
        return float(pos.group(1).replace(',', ''))
    return None


def is_date_line(line):
    """Check if line is a date like 'FEB 17' or a year like '2026'."""
    return bool(re.match(rf'^{MONTH_PAT}\s+\d{{1,2}}$', line, re.IGNORECASE)) or \
           bool(re.match(r'^\d{4}$', line))


def is_amount_line(line):
    """Check if line is a dollar amount (positive or negative in parentheses)."""
    cleaned = line.rstrip(';:. ')
    # Match $1,234.56 or ($1,234.56)
    return bool(re.match(r'^\(?\$?[\d,]+\.\d{2}\)?$', cleaned))


def pair_amounts(amount_lines):
    """Pair amount lines into (transaction_amount, running_balance) tuples.

    Each transaction has an amount followed by a running balance.
    Negative amounts (debits) are in parentheses: ($1,234.56)
    The running balance is always positive.
    """
    pairs = []
    i = 0
    while i < len(amount_lines):
        amt = parse_amount(amount_lines[i])
        bal = None
        # The next line should be the running balance (always positive)
        if i + 1 < len(amount_lines):
            next_val = parse_amount(amount_lines[i + 1])
            if next_val is not None and next_val > 0:
                bal = next_val
                i += 1
        pairs.append((amt, bal))
        i += 1
    return pairs


def parse_dates_from_raw(dates_raw):
    """Convert raw date lines into 'M/D/YYYY' strings."""
    dates = []
    i = 0
    while i < len(dates_raw):
        m = re.match(rf'^{MONTH_PAT}\s+(\d{{1,2}})$', dates_raw[i], re.IGNORECASE)
        if m:
            month, day = m.group(1).upper(), int(m.group(2))
            yr = 2026
            if i + 1 < len(dates_raw) and re.match(r'^\d{4}$', dates_raw[i + 1]):
                yr = int(dates_raw[i + 1])
                i += 1
            dates.append(f"{MONTH_MAP[month]}/{day}/{yr}")
        i += 1
    return dates


def parse_standard_page(text, page_num):
    """Parse a standard page where OCR produces three blocks: dates, descriptions, amounts.

    Works for pages 2-7 of the bank register where the three-column layout
    produces clean block-separated OCR output.
    """
    lines = [l.strip() for l in text.split('\n') if l.strip()]

    skip_patterns = [
        r'^A(?:tlantic)?$', r'^4?\s*Union Bank', r'^Good Morning',
        r'^Old Checking', r'^Last Updated', r'^Current Balance',
        r'^Available Balance', r'^Transactions\s+Details',
        r'^\$[\d,]+\.\d{2}\s+\$[\d,]+\.\d{2}$',
        r'^Date\s+Description', r'^Amount$', r'^Page totals:',
        r'^\d+\s*-\s*\d+\s+of\s+\d', r'^[<>]+$',
    ]
    filtered = []
    for line in lines:
        if any(re.match(p, line, re.IGNORECASE) for p in skip_patterns):
            continue
        filtered.append(line)

    dates_raw = []
    descriptions = []
    amount_lines = []
    state = 'dates'

    for line in filtered:
        if state == 'dates':
            if is_date_line(line):
                dates_raw.append(line)
            elif is_amount_line(line):
                state = 'amounts'
                amount_lines.append(line.rstrip(';:. '))
            else:
                state = 'descriptions'
                if len(line) > 3:
                    descriptions.append(line)
        elif state == 'descriptions':
            if is_amount_line(line):
                state = 'amounts'
                amount_lines.append(line.rstrip(';:. '))
            elif not is_date_line(line) and len(line) > 3:
                descriptions.append(line)
        elif state == 'amounts':
            if is_amount_line(line):
                amount_lines.append(line.rstrip(';:. '))

    dates = parse_dates_from_raw(dates_raw)
    amt_pairs = pair_amounts(amount_lines)

    n = min(len(dates), len(descriptions), len(amt_pairs))
    transactions = []
    for i in range(n):
        amt, bal = amt_pairs[i]
        if amt is not None:
            transactions.append({
                'date': dates[i],
                'page': page_num,
                'description': descriptions[i],
                'amount': amt,
                'balance': bal,
            })
    return transactions


def parse_first_page(img, page_num):
    """Parse the first page which has a header and garbled dates in full OCR.

    Strategy:
    - Crop the date column and OCR separately for clean dates
    - Extract descriptions from full OCR by stripping garbled date prefixes
    - Extract amounts from full OCR normally
    """
    w, h = img.size

    # 1. Get dates from cropped date column
    date_crop = img.crop((0, int(h * 0.42), int(w * 0.15), h))
    date_text = pytesseract.image_to_string(date_crop, config='--psm 4')

    dates = []
    has_pending = False
    dlines = [l.strip() for l in date_text.split('\n') if l.strip()]
    di = 0
    while di < len(dlines):
        dl = dlines[di]
        if re.match(r'^Pending', dl, re.IGNORECASE):
            has_pending = True
            di += 1
            continue
        m = re.match(rf'^{MONTH_PAT}\s+(\d{{1,2}})', dl, re.IGNORECASE)
        if m:
            mo, dy = m.group(1).upper(), int(m.group(2))
            yr = 2026
            if di + 1 < len(dlines):
                ym = re.match(r'^(\d{4})', dlines[di + 1])
                if ym:
                    parsed_yr = int(ym.group(1))
                    # Sanity check: OCR sometimes garbles year digits
                    if 2020 <= parsed_yr <= 2030:
                        yr = parsed_yr
                    di += 1
            dates.append(f"{MONTH_MAP[mo]}/{dy}/{yr}")
        di += 1

    # 2. Get descriptions and amounts from full OCR
    full_text = pytesseract.image_to_string(img)
    lines = [l.strip() for l in full_text.split('\n') if l.strip()]

    descriptions = []
    amount_lines = []

    pending_desc_skipped = False
    for line in lines:
        if is_amount_line(line):
            amount_lines.append(line.rstrip(';:. '))
        elif any(kw in line.upper() for kw in DESC_KEYWORDS):
            # Skip the Pending entry's description (first desc line containing "Pending")
            if has_pending and not pending_desc_skipped and 'pending' in line.lower():
                pending_desc_skipped = True
                continue
            # Strip garbled date prefix (e.g., "se TUITIONEXPRESS..." -> "TUITIONEXPRESS...")
            best_pos = len(line)
            for kw in DESC_KEYWORDS:
                pos = line.upper().find(kw)
                if pos != -1 and pos < best_pos:
                    best_pos = pos
            if best_pos < len(line):
                clean_desc = line[best_pos:].strip()
                if len(clean_desc) > 3:
                    descriptions.append(clean_desc)

    # 3. Pair amounts and handle Pending
    amt_pairs = pair_amounts(amount_lines)

    if has_pending and amt_pairs:
        # Remove the Pending amount (first entry, no balance)
        if amt_pairs[0][1] is None:
            amt_pairs = amt_pairs[1:]
        else:
            # Pending amount was paired with the next amount as "balance"
            # Reconstruct: drop the first value, re-pair everything
            all_vals = []
            if amt_pairs[0][1] is not None:
                all_vals.append(amt_pairs[0][1])
            for a, b in amt_pairs[1:]:
                if a is not None:
                    all_vals.append(a)
                if b is not None:
                    all_vals.append(b)
            amt_pairs = []
            k = 0
            while k < len(all_vals):
                a = all_vals[k]
                b = None
                if k + 1 < len(all_vals) and all_vals[k + 1] > 0:
                    b = all_vals[k + 1]
                    k += 1
                amt_pairs.append((a, b))
                k += 1

    # 4. Build transactions
    n = min(len(dates), len(descriptions), len(amt_pairs))
    transactions = []
    for i in range(n):
        amt, bal = amt_pairs[i]
        if amt is not None:
            transactions.append({
                'date': dates[i],
                'page': page_num,
                'description': descriptions[i],
                'amount': amt,
                'balance': bal,
            })
    return transactions


def parse_last_page(text, page_num):
    """Parse the last page where OCR may put date/desc/amount on the same lines.

    On the last page (few transactions), Tesseract often merges columns into single lines like:
    'DEC 30 FIDELITY INVESTM/GrantPaymt CORNERSTONES INC $2,200.00'
    """
    lines = [l.strip() for l in text.split('\n') if l.strip()]

    # Filter footer
    lines = [l for l in lines if not re.match(r'^Page totals:', l, re.IGNORECASE)
             and not re.match(r'^\d+\s*-\s*\d+\s+of\s+\d', l)]

    # Collect all raw content
    all_text = '\n'.join(lines)

    # Try to find date-description-amount patterns across lines
    # Pattern: "MON DD" on one line, description (possibly with amount) on next, year+balance on next
    transactions = []
    i = 0
    while i < len(lines):
        line = lines[i]

        # Look for a date start: "DEC 30 ..." possibly with description and amount on same line
        date_match = re.match(rf'^{MONTH_PAT}\s+(\d{{1,2}})\s+(.*)', line, re.IGNORECASE)
        if date_match:
            month = date_match.group(1).upper()
            day = int(date_match.group(2))
            rest = date_match.group(3).strip()

            # Extract amount from the rest if present
            amt_in_line = re.search(r'(\(?\$[\d,]+\.\d{2}\)?)\s*[;:.]?\s*$', rest)
            if amt_in_line:
                amount_str = amt_in_line.group(1)
                desc_part = rest[:amt_in_line.start()].strip()
            else:
                amount_str = None
                desc_part = rest

            desc_part = desc_part.strip().rstrip(':;.')

            # Look ahead for description (if not on date line) and year/balance
            year = 2026
            balance = None

            # If description is empty, the next line may have it
            if not desc_part and i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                if any(kw in next_line.upper() for kw in DESC_KEYWORDS):
                    desc_part = next_line.rstrip(':;.')
                    i += 1

            # Next line after description should be year + balance
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                yr_match = re.match(r'^(\d{4})\s*(.*)', next_line)
                if yr_match:
                    year = int(yr_match.group(1))
                    rest2 = yr_match.group(2).strip()
                    bal_match = re.search(r'(\$[\d,]+\.\d{2})', rest2)
                    if bal_match:
                        balance = parse_amount(bal_match.group(1))
                    i += 1

            amount = parse_amount(amount_str) if amount_str else None

            if desc_part and amount is not None:
                # Clean description: find earliest keyword match position
                # but keep the full description if it starts reasonably
                best_pos = len(desc_part)
                for kw in DESC_KEYWORDS:
                    pos = desc_part.upper().find(kw)
                    if pos != -1 and pos < best_pos:
                        best_pos = pos
                if best_pos > 0 and best_pos < len(desc_part):
                    desc_part = desc_part[best_pos:]

                date_str = f"{MONTH_MAP[month]}/{day}/{year}"
                transactions.append({
                    'date': date_str,
                    'page': page_num,
                    'description': desc_part.strip(),
                    'amount': amount,
                    'balance': balance,
                })

        i += 1

    return transactions


def parse_account_info(img):
    """Extract account name from first page header."""
    w, h = img.size
    header_crop = img.crop((0, 0, w, int(h * 0.35)))
    text = pytesseract.image_to_string(header_crop)
    match = re.search(r'((?:Old |New )?Checking Account\s*\*\*\d+)', text)
    return match.group(1).strip() if match else "Bank Register"


def build_excel(transactions, account_name):
    """Build formatted Excel workbook with debits, credits, and running balances."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Bank Register"

    header_fill = PatternFill('solid', fgColor='4472C4')
    header_font = Font(name='Arial', bold=True, size=11, color='FFFFFF')
    data_font = Font(name='Arial', size=10)
    bold_font = Font(name='Arial', bold=True, size=10)
    money_fmt = '#,##0.00'
    thin_border = Border(bottom=Side(style='thin', color='D9D9D9'))

    headers = ['Date', 'Page', 'Description', 'Debits (Out)', 'Credits (In)', 'Balance', 'Status']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', wrap_text=True)

    # Reverse to chronological (oldest first)
    transactions.reverse()

    # Beginning balance = first transaction's balance minus its amount
    if transactions:
        first = transactions[0]
        beginning_balance = round((first['balance'] or 0) - first['amount'], 2)
    else:
        beginning_balance = 0

    ws.cell(row=2, column=1, value='Beginning Balance').font = bold_font
    ws.cell(row=2, column=6, value=beginning_balance).font = bold_font
    ws.cell(row=2, column=6).number_format = money_fmt

    row = 2
    for txn in transactions:
        row += 1
        ws.cell(row=row, column=1, value=txn['date']).font = data_font
        ws.cell(row=row, column=2, value=txn['page']).font = data_font
        ws.cell(row=row, column=2).alignment = Alignment(horizontal='center')
        ws.cell(row=row, column=3, value=txn['description']).font = data_font

        if txn['amount'] < 0:
            ws.cell(row=row, column=4, value=abs(txn['amount'])).number_format = money_fmt
            ws.cell(row=row, column=4).font = data_font
        else:
            ws.cell(row=row, column=5, value=txn['amount']).number_format = money_fmt
            ws.cell(row=row, column=5).font = data_font

        prev = 'F2' if row == 3 else f'F{row-1}'
        ws.cell(row=row, column=6).value = f'={prev}-D{row}+E{row}'
        ws.cell(row=row, column=6).number_format = money_fmt
        ws.cell(row=row, column=6).font = data_font

        ws.cell(row=row, column=7, value='').font = data_font
        for c in range(1, 8):
            ws.cell(row=row, column=c).border = thin_border

    last_data = row

    # TOTALS row
    row += 1
    ws.cell(row=row, column=1, value='TOTALS').font = bold_font
    fill = PatternFill('solid', fgColor='D9E2F3')
    for c in range(1, 8):
        ws.cell(row=row, column=c).fill = fill
    ws.cell(row=row, column=4).value = f'=SUM(D3:D{last_data})'
    ws.cell(row=row, column=4).number_format = money_fmt
    ws.cell(row=row, column=4).font = bold_font
    ws.cell(row=row, column=5).value = f'=SUM(E3:E{last_data})'
    ws.cell(row=row, column=5).number_format = money_fmt
    ws.cell(row=row, column=5).font = bold_font
    ws.cell(row=row, column=6).value = f'=F{last_data}'
    ws.cell(row=row, column=6).number_format = money_fmt
    ws.cell(row=row, column=6).font = bold_font

    # Summary
    row += 2
    ws.cell(row=row, column=1, value='Total items:').font = bold_font
    ws.cell(row=row, column=4, value=len(transactions)).font = bold_font
    row += 1
    ws.cell(row=row, column=1, value='Beginning balance:').font = bold_font
    ws.cell(row=row, column=6, value=beginning_balance).font = bold_font
    ws.cell(row=row, column=6).number_format = money_fmt
    row += 1
    ws.cell(row=row, column=1, value='Ending balance:').font = bold_font
    ws.cell(row=row, column=6).value = f'=F{last_data}'
    ws.cell(row=row, column=6).font = bold_font
    ws.cell(row=row, column=6).number_format = money_fmt

    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 8
    ws.column_dimensions['C'].width = 55
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 16
    ws.column_dimensions['F'].width = 16
    ws.column_dimensions['G'].width = 22
    ws.freeze_panes = 'A2'

    return wb


# --- Streamlit UI ---
uploaded_file = st.file_uploader("Upload Bank Register PDF", type="pdf")

if uploaded_file:
    if st.button("Convert to Excel"):
        with st.spinner("Running OCR on PDF pages... This may take a minute."):
            images = ocr_pdf_to_images(uploaded_file)
        st.info(f"Processed {len(images)} pages via OCR.")

        with st.spinner("Parsing transactions..."):
            account_name = parse_account_info(images[0])
            all_transactions = []

            for page_num, img in enumerate(images, 1):
                text = pytesseract.image_to_string(img)

                if page_num == 1:
                    txns = parse_first_page(img, page_num)
                elif page_num == len(images):
                    txns = parse_last_page(text, page_num)
                else:
                    txns = parse_standard_page(text, page_num)

                all_transactions.extend(txns)

        st.success(f"Found {len(all_transactions)} transactions from '{account_name}'")

        if all_transactions:
            preview = []
            for t in all_transactions[:10]:
                preview.append({
                    'Date': t['date'],
                    'Page': t['page'],
                    'Description': t['description'][:60],
                    'Debit': f"${abs(t['amount']):,.2f}" if t['amount'] < 0 else '',
                    'Credit': f"${t['amount']:,.2f}" if t['amount'] >= 0 else '',
                    'Balance': f"${t['balance']:,.2f}" if t['balance'] else '',
                })
            st.write("**Preview (first 10 transactions, newest first):**")
            st.table(preview)

        with st.spinner("Building Excel file..."):
            wb = build_excel(all_transactions, account_name)
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

        st.download_button(
            label="Download Excel File",
            data=output.getvalue(),
            file_name=f"Bank_Register_{account_name.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.officedocument",
        )
