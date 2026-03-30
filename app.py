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
    if text is None:
        return None
    text = text.strip().rstrip(';:., ')
    neg = re.match(r'^\(\$?([\d,]+\.\d{2})\)$', text)
    if neg:
        return -float(neg.group(1).replace(',', ''))
    pos = re.match(r'^\$?([\d,]+\.\d{2})$', text)
    if pos:
        return float(pos.group(1).replace(',', ''))
    return None


def extract_amounts_from_text(text):
    """Extract all dollar amounts from a text string, returning (amount_str, start_pos) pairs."""
    pattern = r'(\(?\$[\d,]+\.\d{2}\)?)'
    return [(m.group(1), m.start()) for m in re.finditer(pattern, text)]


def is_date_line(line):
    """Check if line is a date like 'FEB 17' or a year like '2026'."""
    return bool(re.match(rf'^{MONTH_PAT}\s+\d{{1,2}}$', line, re.IGNORECASE)) or \
           bool(re.match(r'^\d{4}$', line))


def is_amount_line(line):
    """Check if line is a dollar amount (positive or negative in parentheses)."""
    cleaned = line.strip().rstrip(';:., ')
    return bool(re.match(r'^\(?\$?[\d,]+\.\d{2}\)?$', cleaned))


def is_header_or_footer(line):
    """Check if a line is a header, footer, or other non-transaction content."""
    skip_patterns = [
        r'^A(?:tlantic)?$', r'^4?\s*Union Bank', r'^Good Morning',
        r'^Old Checking', r'^Last Updated', r'^Current Balance',
        r'^Available Balance', r'^Transactions\s+Details',
        r'^\$[\d,]+\.\d{2}\s+\$[\d,]+\.\d{2}$',
        r'^Date\s+Description', r'^Amount$', r'^Page totals:',
        r'^\d+\s*-\s*\d+\s+of\s+\d', r'^[<>]+$', r'^Pending\b',
        r'^Details\s*&\s*Settings',
    ]
    return any(re.match(p, line, re.IGNORECASE) for p in skip_patterns)


# ---------------------------------------------------------------------------
# Parser for "three-block" pages (standard pages where OCR separates columns)
# OCR output: all dates first, then all descriptions, then all amounts.
# ---------------------------------------------------------------------------

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


def pair_amounts(amount_lines):
    """Pair consecutive amount lines into (transaction_amount, running_balance) tuples.

    Each transaction has an amount followed by a running balance.
    The running balance is always positive. Debits are in parentheses.
    """
    pairs = []
    i = 0
    while i < len(amount_lines):
        amt = parse_amount(amount_lines[i])
        bal = None
        if i + 1 < len(amount_lines):
            next_val = parse_amount(amount_lines[i + 1])
            if next_val is not None and next_val > 0:
                bal = next_val
                i += 1
        pairs.append((amt, bal))
        i += 1
    return pairs


def parse_block_page(text, page_num):
    """Parse a page where OCR produces three clean blocks: dates, descriptions, amounts.

    This is the standard format for middle pages of the bank register.
    """
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    filtered = [l for l in lines if not is_header_or_footer(l)]

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
                amount_lines.append(line.rstrip(';:., '))
            else:
                state = 'descriptions'
                if len(line) > 2:
                    descriptions.append(line)
        elif state == 'descriptions':
            if is_amount_line(line):
                state = 'amounts'
                amount_lines.append(line.rstrip(';:., '))
            elif not is_date_line(line) and len(line) > 2:
                descriptions.append(line)
        elif state == 'amounts':
            if is_amount_line(line):
                amount_lines.append(line.rstrip(';:., '))

    dates = parse_dates_from_raw(dates_raw)
    amt_pairs = pair_amounts(amount_lines)

    # Use the minimum count but warn if they don't match
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


def is_block_format(text):
    """Detect if OCR output is in three-block format (dates, then descriptions, then amounts).

    In block format, the first non-header lines are all dates/years,
    then switch to non-date/non-amount text, then switch to amounts.
    """
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    filtered = [l for l in lines if not is_header_or_footer(l)]

    if len(filtered) < 6:
        return False

    # Check if first several lines are dates
    date_count = 0
    for line in filtered:
        if is_date_line(line):
            date_count += 1
        else:
            break

    # Block format has at least 4 date/year lines at the start (2 transactions minimum)
    return date_count >= 4


# ---------------------------------------------------------------------------
# Parser for "merged" pages (first/last pages where OCR mixes columns)
# Each transaction spans 2 lines:
#   Line 1: "MON DD [description] [amount] [noise]"
#   Line 2: "YYYY [description_cont] [balance] [noise]"
# ---------------------------------------------------------------------------

def parse_merged_page(text, page_num):
    """Parse a page where OCR merges date/description/amount on the same lines.

    Works for pages where Tesseract produces lines like:
      'MAR 27 COVA/VENDORPAYM Cornerstones, Inc. $22,482.54'
      '2026 $25,692.37'
    or:
      'JAN > COMMONWEALTH OF/ECC PMTS LAUREL LEARNING CENTER $17,331.00;'
      '2026 $23,578.28'
    """
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    filtered = [l for l in lines if not is_header_or_footer(l)]

    transactions = []
    i = 0
    while i < len(filtered):
        line = filtered[i]

        # Look for a line starting with a month name (date line)
        # Handle garbled OCR: "JAN >" might be "JAN 5", "MAR 27°" etc.
        date_match = re.match(
            rf'^{MONTH_PAT}\s+(\S{{1,2}})',
            line, re.IGNORECASE
        )
        if not date_match:
            i += 1
            continue

        month = date_match.group(1).upper()
        day_str = date_match.group(2).strip()

        # Handle garbled day: OCR often misreads digits as symbols
        OCR_DAY_FIXES = {'>': '5', '|': '1', 'l': '1', 'O': '0',
                         'o': '0', 'S': '5', 's': '5', 'Z': '2',
                         'z': '2', 'B': '8', 'G': '6', 'q': '9'}
        cleaned_day = ''.join(OCR_DAY_FIXES.get(c, c) for c in day_str)
        try:
            day = int(cleaned_day)
        except ValueError:
            i += 1
            continue

        rest_of_line = line[date_match.end():].strip()

        # Extract amounts from the rest of the date line
        amounts_in_line = extract_amounts_from_text(rest_of_line)

        # Get description: everything before the first amount
        if amounts_in_line:
            first_amt_pos = amounts_in_line[0][1]
            desc_part = rest_of_line[:first_amt_pos].strip().rstrip(':;., ')
            txn_amount_str = amounts_in_line[0][0]
        else:
            desc_part = rest_of_line.strip().rstrip(':;., ')
            txn_amount_str = None

        # Clean description: remove leading symbols like "_ ", "= ", "°"
        desc_part = re.sub(r'^[_=°®\s]+', '', desc_part).strip()

        # Handle 3-line format: date+amount, description, year+balance
        # e.g.: "DEC 30 $2,200.00" / "FIDELITY INVESTM/..." / "2025 $3,667.53"
        if not desc_part and i + 1 < len(filtered):
            peek = filtered[i + 1]
            # If next line is not a date, not a year, and not a header → it's a description
            if not re.match(rf'^{MONTH_PAT}\s+', peek, re.IGNORECASE) and \
               not re.match(r'^[25]\d{3}\b', peek) and \
               not is_header_or_footer(peek) and \
               not is_amount_line(peek):
                desc_part = re.sub(r'^[_=°®©@&\s]+', '', peek).strip().rstrip(':;., ')
                i += 1  # consumed the description line

        # Look at next line for year and/or balance
        year = 2026
        balance = None

        if i + 1 < len(filtered):
            next_line = filtered[i + 1]
            yr_match = re.match(r'^[25]\d{3}\b', next_line)
            if yr_match:
                parsed_yr = int(yr_match.group(0))
                if 2020 <= parsed_yr <= 2030:
                    year = parsed_yr

                year_rest = next_line[yr_match.end():].strip()

                # The year line might also contain description and/or balance
                amounts_in_year = extract_amounts_from_text(year_rest)

                if amounts_in_year:
                    # Text before the amount on the year line could be description
                    year_desc = year_rest[:amounts_in_year[0][1]].strip().rstrip(':;., ')
                    year_desc = re.sub(r'^[_=°®\s]+', '', year_desc).strip()

                    if not desc_part and year_desc:
                        desc_part = year_desc
                    elif year_desc and not any(c.isalpha() for c in desc_part):
                        desc_part = year_desc

                    # Balance is typically the last amount on the year line
                    balance = parse_amount(amounts_in_year[-1][0])

                    # If there's no txn amount from the date line,
                    # and the year line has 2 amounts, first is amount, second is balance
                    if txn_amount_str is None and len(amounts_in_year) >= 2:
                        txn_amount_str = amounts_in_year[0][0]
                        balance = parse_amount(amounts_in_year[1][0])
                elif not desc_part:
                    # Year line has no amounts - might be pure description
                    year_desc = year_rest.strip().rstrip(':;., ')
                    year_desc = re.sub(r'^[_=°®\s]+', '', year_desc).strip()
                    if year_desc:
                        desc_part = year_desc

                i += 1  # consumed the year line

        txn_amount = parse_amount(txn_amount_str) if txn_amount_str else None
        date_str = f"{MONTH_MAP[month]}/{day}/{year}"

        if txn_amount is not None and desc_part:
            transactions.append({
                'date': date_str,
                'page': page_num,
                'description': desc_part,
                'amount': txn_amount,
                'balance': balance,
            })
        elif txn_amount is not None:
            # Transaction with amount but no description (e.g., DEPOSIT)
            transactions.append({
                'date': date_str,
                'page': page_num,
                'description': 'DEPOSIT',
                'amount': txn_amount,
                'balance': balance,
            })

        i += 1

    return transactions


# ---------------------------------------------------------------------------
# Unified page parser: auto-detect format and dispatch
# ---------------------------------------------------------------------------

def parse_page(img, page_num, total_pages):
    """Parse a single page, auto-detecting whether it uses block or merged format.

    If default OCR produces 0 transactions (common on page 1 where dates get garbled),
    retry with --psm 4 which often produces cleaner merged-format output.
    """
    text = pytesseract.image_to_string(img)

    if is_block_format(text):
        txns = parse_block_page(text, page_num)
    else:
        txns = parse_merged_page(text, page_num)

    # Fallback: retry with --psm 4 if no transactions found
    if not txns:
        text_psm4 = pytesseract.image_to_string(img, config='--psm 4')
        txns = parse_merged_page(text_psm4, page_num)

    return txns


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
    if transactions and transactions[0]['balance'] is not None:
        first = transactions[0]
        beginning_balance = round(first['balance'] - first['amount'], 2)
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
                txns = parse_page(img, page_num, len(images))
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
