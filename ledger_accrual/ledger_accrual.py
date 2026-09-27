import os
from datetime import datetime, timedelta
import calendar
import warnings
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

# ==========================================
# 1. DATE & LABEL UTILITIES
# ==========================================

def get_period_end(date, frequency):
    if frequency == 'M':
        _, last_day = calendar.monthrange(date.year, date.month)
        return datetime(date.year, date.month, last_day)
    elif frequency == 'Q':
        if 1 <= date.month <= 3: return datetime(date.year, 3, 31)
        elif 4 <= date.month <= 6: return datetime(date.year, 6, 30)
        elif 7 <= date.month <= 9: return datetime(date.year, 9, 30)
        else: return datetime(date.year, 12, 31)
    elif frequency == 'Y':
        if date.month <= 3: return datetime(date.year, 3, 31)
        else: return datetime(date.year + 1, 3, 31)
    else:
        raise ValueError(f"Unknown frequency: {frequency}")

def get_interest_label(period_end, frequency):
    if frequency == 'M':
        return f"Monthly Auto Interest - {period_end.strftime('%B %Y')}"
    elif frequency == 'Q':
        q_str = {3: "Jan-Mar", 6: "Apr-Jun", 9: "Jul-Sep", 12: "Oct-Dec"}[period_end.month]
        return f"Quarterly Auto Interest - {q_str} {period_end.year}"
    elif frequency == 'Y':
        return f"Yearly Auto Interest - FY {period_end.year - 1}-{str(period_end.year)[-2:]}"

def format_rate_display(rate_fraction):
    """Turn a rate fraction (0.12) into a compact percent string ('12' or '12.5')."""
    pct = rate_fraction * 100.0
    if abs(pct - round(pct)) < 1e-9:
        return str(int(round(pct)))
    return f"{pct:g}"

def build_rate_suffix(segments):
    """
    Build the rate suffix for an interest entry from the applied rate segments.

    `segments` is an ordered list of dicts {'rate', 'from', 'to'} describing the
    actual date span each rate was applied within this interest period.

    - Single rate:    ' @16%'
    - Multiple rates: ' @16% (01/07/2025 to 15/09/2025) // 18% (16/09/2025 to 20/09/2025)'
    """
    if not segments:
        return ""

    if len(segments) == 1:
        return " @" + format_rate_display(segments[0]['rate']) + "%"

    parts = []
    for s in segments:
        parts.append(
            f"{format_rate_display(s['rate'])}% "
            f"({s['from'].strftime('%d/%m/%Y')} to {s['to'].strftime('%d/%m/%Y')})"
        )
    return " @" + " // ".join(parts)

# ==========================================
# 2. DATA PROCESSING
# ==========================================

def parse_date(val):
    if isinstance(val, datetime):
        return val
    if isinstance(val, str):
        val = val.strip()
        date_formats = [
            '%d/%m/%Y', '%d-%b-%y', '%d-%b-%Y', 
            '%d-%m-%Y', '%Y-%m-%d', '%d/%m/%y'
        ]
        for fmt in date_formats:
            try:
                return datetime.strptime(val, fmt)
            except ValueError:
                continue
        print(f"[!] WARNING: Could not parse date string: '{val}'. Ensure it is a valid date format.")
    return None

def safe_float(val):
    if val is None or str(val).strip() == "":
        return 0.0
    try:
        return float(val)
    except ValueError:
        return 0.0

def create_base_row(headers):
    return {h: None for h in headers}

def parse_rate_file(rate_file):
    """
    Parse rate.txt located next to the ledger file.

    Expected format:
        frequency: Q
        from: 01/04/2023, to: 30/09/2023, rate: 12
        from: 01/10/2023, to: 31/03/2024, rate: 15

    Returns (freq, ranges) where ranges is a list of dicts:
        {'from': datetime, 'to': datetime, 'rate': fraction}
    sorted by 'from'. Rate is stored as a fraction (12 -> 0.12).
    """
    if not os.path.exists(rate_file):
        raise ValueError(f"Required rate file not found: {rate_file}")

    freq = None
    ranges = []

    with open(rate_file, 'r', encoding='utf-8') as f:
        for line_no, raw in enumerate(f, start=1):
            line = raw.strip()
            if not line or line.startswith('#'):
                continue

            low = line.lower()
            if low.startswith('frequency:'):
                freq = line.split(':', 1)[1].strip().upper()
                continue

            # Expect a from/to/rate triplet line
            parts = {}
            for chunk in line.split(','):
                if ':' not in chunk:
                    continue
                key, val = chunk.split(':', 1)
                parts[key.strip().lower()] = val.strip()

            if not {'from', 'to', 'rate'}.issubset(parts.keys()):
                raise ValueError(
                    f"rate.txt line {line_no} is malformed: '{line}'. "
                    f"Expected 'from: DD/MM/YYYY, to: DD/MM/YYYY, rate: X'."
                )

            from_dt = parse_date(parts['from'])
            to_dt = parse_date(parts['to'])
            if from_dt is None or to_dt is None:
                raise ValueError(f"rate.txt line {line_no} has an unparseable date: '{line}'")
            if to_dt < from_dt:
                raise ValueError(f"rate.txt line {line_no}: 'to' date is before 'from' date.")

            try:
                rate_fraction = float(parts['rate']) / 100.0
            except ValueError:
                raise ValueError(f"rate.txt line {line_no}: rate '{parts['rate']}' is not a number.")

            ranges.append({'from': from_dt, 'to': to_dt, 'rate': rate_fraction})

    if freq is None:
        raise ValueError("rate.txt is missing the 'frequency:' line (expected M, Q, or Y).")
    if freq not in ('M', 'Q', 'Y'):
        raise ValueError(f"rate.txt frequency must be M, Q, or Y, but got '{freq}'.")
    if not ranges:
        raise ValueError("rate.txt does not define any 'from/to/rate' ranges.")

    ranges.sort(key=lambda r: r['from'])

    # Detect overlapping ranges (ambiguous rate for a day)
    for prev, curr in zip(ranges, ranges[1:]):
        if curr['from'] <= prev['to']:
            raise ValueError(
                f"rate.txt has overlapping date ranges: "
                f"{prev['from'].strftime('%d/%m/%Y')}-{prev['to'].strftime('%d/%m/%Y')} and "
                f"{curr['from'].strftime('%d/%m/%Y')}-{curr['to'].strftime('%d/%m/%Y')}."
            )

    return freq, ranges

def range_for_day(day, ranges):
    """Return the rate range dict covering `day`, or None if uncovered. Inclusive of from/to."""
    for r in ranges:
        if r['from'] <= day <= r['to']:
            return r
    return None

def accrue_interest(balance, start_date, end_date, ranges):
    """
    Accrue interest on `balance` for each day in [start_date, end_date), using the
    per-day rate from `ranges`. Each day contributes balance * daily_rate.

    Returns (interest, segments) where `segments` is an ordered list of
        {'rate': fraction, 'from': datetime, 'to': datetime}
    describing the actual days a rate applied within [start_date, end_date).
    'from'/'to' are the first and last day (inclusive) the rate was active in this
    window. Raises ValueError if any day is not covered by a range.
    """
    interest = 0.0
    segments = []  # list of {'rate', 'from', 'to'}, contiguous per rate change
    day = start_date
    while day < end_date:
        rng = range_for_day(day, ranges)
        if rng is None:
            raise ValueError(
                f"No interest rate defined in rate.txt for date {day.strftime('%d/%m/%Y')}."
            )
        interest += balance * rng['rate'] * (1.0 / 365.0)

        if segments and segments[-1]['rate'] == rng['rate']:
            segments[-1]['to'] = day
        else:
            segments.append({'rate': rng['rate'], 'from': day, 'to': day})
        day += timedelta(days=1)
    return interest, segments

def process_ledger_math(headers, rows, ranges, freq):
    new_rows = []
    current_balance = 0.0
    accumulated_interest = 0.0
    total_auto_interest = 0.0
    period_segments = []  # rate segments (with clamped dates) applied within the current period
    last_date = None

    def add_period_segments(segs):
        # Segments arrive in chronological order; merge with the running list,
        # extending the last segment when the rate matches (contiguous same-rate span).
        for s in segs:
            if period_segments and period_segments[-1]['rate'] == s['rate']:
                period_segments[-1]['to'] = s['to']
            else:
                period_segments.append(dict(s))

    def post_interest(period_end):
        nonlocal current_balance, accumulated_interest, total_auto_interest, period_segments
        interest_to_post = round(accumulated_interest, 0)
        current_balance += interest_to_post
        total_auto_interest += interest_to_post

        interest_row = create_base_row(headers)
        interest_row['DATE'] = period_end + timedelta(days=1)
        interest_row['PARTICULAR'] = get_interest_label(period_end, freq) + build_rate_suffix(period_segments)
        interest_row['DEBIT'] = interest_to_post
        interest_row['BALANCE'] = current_balance
        new_rows.append(interest_row)

        accumulated_interest = 0.0
        period_segments = []

    for row in rows:
        current_date = row['DATE']
        if current_date is None: continue
        
        if last_date is None:
            last_date = current_date
            
        period_end = get_period_end(last_date, freq)
        next_period_start = period_end + timedelta(days=1)
        
        while current_date >= next_period_start:
            interest, used = accrue_interest(current_balance, last_date, next_period_start, ranges)
            accumulated_interest += interest
            add_period_segments(used)
            post_interest(period_end)

            last_date = next_period_start
            period_end = get_period_end(last_date, freq)
            next_period_start = period_end + timedelta(days=1)

        if current_date > last_date:
            interest, used = accrue_interest(current_balance, last_date, current_date, ranges)
            accumulated_interest += interest
            add_period_segments(used)
            last_date = current_date
        
        debit_val = safe_float(row.get('DEBIT'))
        credit_val = safe_float(row.get('CREDIT'))
        
        # INVERTED LOGIC: Balance = Balance + Credit - Debit
        current_balance += credit_val
        current_balance -= debit_val
        
        row['BALANCE'] = current_balance
        new_rows.append(row)

    # Trailing accrual: roll forward through every remaining period boundary up to
    # the end of the rate coverage defined in rate.txt, posting interest for each.
    if last_date is not None and ranges:
        coverage_end = ranges[-1]['to']
        period_end = get_period_end(last_date, freq)

        # Post interest for each period boundary up to and including the period
        # that contains the end of the rate coverage defined in rate.txt.
        while last_date <= coverage_end:
            # Accrue up to the earlier of this period's end or the coverage end
            # (both inclusive), so the final partial period is captured.
            accrue_to = min(period_end, coverage_end) + timedelta(days=1)
            if accrue_to > last_date:
                interest, used = accrue_interest(current_balance, last_date, accrue_to, ranges)
                accumulated_interest += interest
                add_period_segments(used)

            post_interest(period_end)

            last_date = period_end + timedelta(days=1)
            period_end = get_period_end(last_date, freq)

    data_rows_count = len(new_rows)
    new_rows.append(create_base_row(headers)) # Empty Row
    
    total_row = create_base_row(headers)
    total_row['PARTICULAR'] = "Total Interest"
    total_row['DEBIT'] = total_auto_interest
    new_rows.append(total_row)

    return new_rows, data_rows_count


# ==========================================
# 3. FILE I/O & FORMATTING
# ==========================================

def get_unique_filepath(output_dir, base_name, freq, rate_display):
    os.makedirs(output_dir, exist_ok=True)
    clean_rate = rate_display.replace('.', '_')
    proposed_name = f"{base_name}_{freq.lower()}_{clean_rate}pct.xlsx"
    file_path = os.path.join(output_dir, proposed_name)
    counter = 1
    while os.path.exists(file_path):
        file_path = os.path.join(output_dir, f"{base_name}_{freq.lower()}_{clean_rate}pct_{counter}.xlsx")
        counter += 1
    return file_path

def process_file(input_file):
    print(f"\n--- Reading {input_file} ---")
    wb = load_workbook(input_file, data_only=True)
    ws = wb.active
    
    # 1. Extract Headers
    headers = [cell.value for cell in ws[1]]
    if len(headers) < 5:
        raise ValueError("File does not have the expected ledger columns (DATE, PARTICULAR, DEBIT, CREDIT, BALANCE).")

    # Load frequency + rate ranges from rate.txt located next to the input file
    rate_file = os.path.join(os.path.dirname(os.path.abspath(input_file)), 'rate.txt')
    freq, ranges = parse_rate_file(rate_file)

    rate_display = "-".join(format_rate_display(r['rate']) for r in ranges)
    print(f"Loaded rate.txt: Frequency={freq}, {len(ranges)} rate range(s)")
    for r in ranges:
        print(f"    {r['from'].strftime('%d/%m/%Y')} -> {r['to'].strftime('%d/%m/%Y')} @ {format_rate_display(r['rate'])}%")

    # 2. Extract Data & Sort chronologically
    parsed_rows = []
    for row_idx, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if all(v is None for v in r): continue
        row_dict = {headers[i]: r[i] for i in range(len(headers))}
        
        # Check for existing manual interest
        part_str = str(row_dict.get('PARTICULAR', '')).lower()
        if 'interest' in part_str:
            print(f"[*] WARNING: Existing interest entry found at Row {row_idx}. Included in calculation.")
            
        date_val = parse_date(row_dict.get('DATE'))
        if date_val:
            row_dict['DATE'] = date_val
            parsed_rows.append(row_dict)

    parsed_rows.sort(key=lambda x: x['DATE'])
    final_rows, data_rows_count = process_ledger_math(headers, parsed_rows, ranges, freq)
    
    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = 'Ledger'
    
    out_ws.append(headers)
    for row_dict in final_rows:
        out_ws.append([row_dict[h] for h in headers])
        
    bold_font = Font(bold=True)
    thin_bottom = Border(bottom=Side(style='thin'))
    wrap_align = Alignment(wrap_text=True, vertical='top')

    # Cap the PARTICULAR column width so wide multi-rate labels wrap instead of
    # stretching the sheet beyond A4 width.
    PARTICULAR_MAX_WIDTH = 45

    date_col = headers.index('DATE') + 1 if 'DATE' in headers else 1
    debit_col = headers.index('DEBIT') + 1 if 'DEBIT' in headers else 3
    credit_col = headers.index('CREDIT') + 1 if 'CREDIT' in headers else 4
    balance_col = headers.index('BALANCE') + 1 if 'BALANCE' in headers else 5
    particular_col = headers.index('PARTICULAR') + 1 if 'PARTICULAR' in headers else 2

    last_ledger_row_idx = data_rows_count + 1
    
    for row_idx, row_dict in enumerate(final_rows, start=2):
        part_val = row_dict.get('PARTICULAR', '')
        
        if row_idx == last_ledger_row_idx:
            for col_idx in range(1, len(headers) + 1):
                out_ws.cell(row=row_idx, column=col_idx).border = thin_bottom
                
        if part_val and ('Auto Interest' in str(part_val) or 'Total Interest' in str(part_val)):
            for col_idx in range(1, len(headers) + 1):
                out_ws.cell(row=row_idx, column=col_idx).font = bold_font
                
        if row_dict.get('DATE'):
            out_ws.cell(row=row_idx, column=date_col).number_format = 'DD/MM/YYYY'
            
        for col_idx in (debit_col, credit_col, balance_col):
            cell = out_ws.cell(row=row_idx, column=col_idx)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00'

        # Wrap the PARTICULAR cell so long labels stay within the capped width.
        out_ws.cell(row=row_idx, column=particular_col).alignment = wrap_align

    for col_idx, col_name in enumerate(headers, start=1):
        col_letter = get_column_letter(col_idx)
        max_length = len(str(col_name))
        for row_dict in final_rows:
            val = row_dict.get(col_name)
            if val is not None:
                if isinstance(val, datetime): max_length = max(max_length, 10)
                else: max_length = max(max_length, len(str(val)))
        width = max_length + 2
        if col_idx == particular_col:
            width = min(width, PARTICULAR_MAX_WIDTH)
        out_ws.column_dimensions[col_letter].width = width

    # Fit the whole table onto A4 width when printed.
    out_ws.page_setup.orientation = 'portrait'
    out_ws.page_setup.paperSize = out_ws.PAPERSIZE_A4
    out_ws.page_setup.fitToWidth = 1
    out_ws.page_setup.fitToHeight = 0
    out_ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    base_name = os.path.splitext(os.path.basename(input_file))[0]
    output_dir = os.path.join(os.getcwd(), 'output')
    output_file = get_unique_filepath(output_dir, base_name, freq, rate_display)
    out_wb.save(output_file)
    
    print(f"[SUCCESS] Processing complete!")
    print(f"Saved formatted output to: {output_file}")

def main():
    print("========================================")
    print("      Ledger Auto-Accrual Processor     ")
    print("========================================\n")
    try:
        current_dir = os.getcwd()
        print(f"Scanning directory: {current_dir}")
        
        excel_files = [f for f in os.listdir(current_dir) if (f.endswith('.xlsx') or f.endswith('.xls')) and not f.startswith('~$')]
        
        if not excel_files:
            print("\n[ERROR] No Excel files (.xlsx or .xls) found in the current directory.")
            return

        print(f"Found {len(excel_files)} Excel file(s).")
        for file in excel_files:
            try:
                process_file(file)
            except Exception as e:
                print(f"\n[ERROR] Failed to process '{file}': {str(e)}")
                print("Skipping to next file (if any)...\n")
                
    except Exception as e:
        print(f"\n[FATAL ERROR] An unexpected error occurred:\n{str(e)}")
    finally:
        print("\n========================================")
        input("Press ENTER to exit...")

if __name__ == "__main__":
    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
    main()