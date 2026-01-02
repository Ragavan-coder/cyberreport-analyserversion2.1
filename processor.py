import re
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ============================================
# CONFIG
# ============================================
BANK_MASTER = {
    "CANARA": "Canara Bank",
    "STATE BANK": "State Bank of India",
    "SBI": "State Bank of India",
    "HDFC": "HDFC Bank",
    "ICICI": "ICICI Bank",
    "AXIS": "Axis Bank",
    "PNB": "Punjab National Bank",
    "FEDERAL": "Federal Bank",
    "UNION": "Union Bank of India",
    "INDUSIND": "IndusInd Bank",
    "UCO": "UCO Bank",
    "CENTRAL": "Central Bank of India",
    "AU": "AU Bank",
    "FINO": "Fino Payments Bank",
    "GOOGLE PAY": "Google Pay",
    "PHONEPE": "PhonePe"
}

# ============================================
# TEXT EXTRACTION
# ============================================
def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        print(f"Error reading PDF: {e}")
    return text

# ============================================
# MAIN FIELDS
# ============================================
def extract_main_fields(text):
    fields = {}
    patterns = {
        "Complaint ID": r'(?:Acknowledgement Number|Complaint ID|Reference ID)\s*:?\s*(\d{10,})',
        "Date Filed": r'(?:Complaint Date|Date of Complaint|Complaint Received|Filed Date)\s*:?\s*([\d\-/]+)',
        "Date Accepted": r'(?:Date Accepted|Accepted Date|Acceptance Date|Registered Date)\s*:?\s*([\d\-/]+)',
        "Time Accepted": r'(?:Time|Registered at|Accepted at)\s*:?\s*([\d\:APMapm ]+)',
        "Complainant Name": r'(?:Name|Complainant Name|Victim Name)\s*:?\s*([A-Za-z\s\*]+?)(?:\n|Email|Contact)',
        "Email": r'(?:Email|E-mail|Email Address)\s*:?\s*([a-zA-Z0-9@\.\*]+)',
        "Mobile Number": r'(?:Mobile|Phone|Contact Number|Mobile Number)\s*:?\s*(\d+)',
        "State": r'(?:State)\s*:?\s*([A-Za-z\s]+?)(?:\n|District)',
        "District": r'(?:District)\s*:?\s*([A-Za-z\s/]+?)(?:\n|Address|City)',
        "Cybercrime Type": r'(?:Category|Crime Type|Category of complaint)\s*:?\s*([A-Za-z\s]+?)(?:\n|Sub)',
        "Sub-Category": r'(?:Sub-Category|Sub Category|Subcategory)\s*:?\s*([A-Za-z\s]+?)(?:\n|IPC|Platform)',
        "Platform": r'(?:Platform|Channel|Mode)\s*:?\s*([A-Za-z\s]+?)(?:\n|Amount|Bank)',
        "Source Bank": r'(?:Bank Name|Source Bank|Account Bank)\s*:?\s*([A-Za-z\s]+?)(?:\n|Account)',
        "Amount Lost": r'(?:Total|Amount|Total Amount|Fraudulent Amount)\s*:?\s*(?:₹|Rs\.?)[\s]*([\d,\.]+)',
        "Transaction Count": r'(?:Number of Transactions|Transaction Count|Total Transactions)\s*:?\s*(\d+)',
        "Date Range": r'(?:Date Range|Fraud Period|Between)\s*:?\s*([\d\-/]+)\s*(?:to|and|-)\s*([\d\-/]+)'
    }

    for field, pattern in patterns.items():
        match = re.search(pattern, text, re.I)
        if match:
            if field == "Date Range":
                fields[field] = f"{match.group(1)} to {match.group(2)}"
            elif field == "Amount Lost":
                amount_str = match.group(1).replace(",", "").replace(".", "")
                fields[field] = f"₹{int(amount_str):,}" if amount_str.isdigit() else ""
            else:
                fields[field] = match.group(1).strip()
        else:
            fields[field] = ""  # Empty if not found

    # Status fields
    text_upper = text.upper()
    fields["Complaint Status"] = "✅ ACCEPTED" if "ACCEPTED" in text_upper else "❌ REJECTED" if "REJECTED" in text_upper else "⏳ PENDING"
    fields["FIR Status"] = "✅ FILED" if "FIR FILED" in text_upper or "FIR NUMBER" in text_upper else "⏳ NOT FILED"
    fields["Investigation Status"] = "✅ CLOSED" if "CLOSED" in text_upper else "🔄 ONGOING" if "INVESTIGATING" in text_upper or "ONGOING" in text_upper else "⏳ NOT STARTED"

    return fields

# ============================================
# TRANSACTIONS
# ============================================
def normalize_bank(text):
    for key, bank in BANK_MASTER.items():
        if key in text.upper():
            return bank
    return "Unknown Bank"

def extract_transactions(text):
    transactions = []
    pattern = re.compile(
        r'(Canara|State Bank|SBI|HDFC|ICICI|Axis|Federal|PNB|Union|IndusInd|UCO|Central|AU|Fino|Google Pay|PhonePe).*?([\d,]{1,}).*?([\d]{1,2}[-/][\d]{1,2}[-/][\d]{4})',
        re.I | re.DOTALL
    )
    txn_num = 1
    seen = set()
    for match in pattern.finditer(text):
        bank = normalize_bank(match.group(1))
        amount_str = match.group(2).replace(",", "")
        try:
            amount = int(amount_str)
        except:
            amount = 0
        date = match.group(3)
        if (date, amount, bank) not in seen and amount > 0:
            transactions.append({
                "Transaction #": txn_num,
                "Date": date,
                "Amount": f"₹{amount:,}",
                "Bank": bank,
                "Status": "Processed"
            })
            seen.add((date, amount, bank))
            txn_num += 1
    return transactions

# ============================================
# DAILY BREAKDOWN
# ============================================
def extract_daily_breakdown(text, transactions):
    daily_data = {}
    for txn in transactions:
        date = txn['Date']
        amount_str = txn['Amount'].replace("₹","").replace(",","")
        try:
            amount = int(amount_str)
        except:
            continue
        if date not in daily_data:
            daily_data[date] = {"total": 0, "count": 0, "banks": set()}
        daily_data[date]["total"] += amount
        daily_data[date]["count"] += 1
        daily_data[date]["banks"].add(txn['Bank'])

    daily_list = []
    for date in sorted(daily_data.keys()):
        data = daily_data[date]
        count = data["count"] if data["count"] > 0 else 1
        daily_list.append({
            "Date": date,
            "Daily Total": f"₹{data['total']:,}",
            "Transaction Count": data["count"],
            "Average per Txn": f"₹{data['total']//count:,}",
            "Banks Involved": ", ".join(list(data["banks"])[:3])
        })
    return daily_list

# ============================================
# DESTINATION BANKS
# ============================================
def extract_destination_banks(text, transactions):
    bank_amounts = {}
    bank_counts = {}
    for txn in transactions:
        bank = txn['Bank']
        amount_str = txn['Amount'].replace("₹","").replace(",","")
        try:
            amount = int(amount_str)
        except:
            amount = 0
        bank_amounts[bank] = bank_amounts.get(bank,0)+amount
        bank_counts[bank] = bank_counts.get(bank,0)+1

    bank_list = []
    total_amount = sum(bank_amounts.values())
    for bank, amount in sorted(bank_amounts.items(), key=lambda x: x[1], reverse=True):
        bank_list.append({
            "Bank/Service": bank,
            "Amount": f"₹{amount:,}" if amount else "",
            "Transfer Count": bank_counts.get(bank,0),
            "% of Total": f"{(amount/total_amount*100):.1f}%" if total_amount else "",
            "Status": "Processed",
            "Recovery Action": ""
        })
    return bank_list

# ============================================
# EXCEL FORMATTING
# ============================================
def format_excel_sheet(ws, header_row=1):
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    border_style = Border(left=Side(style='thin'), right=Side(style='thin'),
                          top=Side(style='thin'), bottom=Side(style='thin'))

    for cell in ws[header_row]:
        if cell.value:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_style

    for row in ws.iter_rows(min_row=header_row+1):
        for cell in row:
            cell.border = border_style
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    for column in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in column)
        column_letter = column[0].column_letter
        ws.column_dimensions[column_letter].width = min(max_length+2, 50)

    ws.freeze_panes = f"A{header_row+1}"

# ============================================
# MAIN PROCESS FUNCTION
# ============================================
def process_pdf(pdf_path, output_excel):
    text = extract_text_from_pdf(pdf_path)
    main_fields = extract_main_fields(text)
    transactions = extract_transactions(text)
    daily_breakdown = extract_daily_breakdown(text, transactions)
    destination_banks = extract_destination_banks(text, transactions)

    wb = Workbook()
    wb.remove(wb.active)

    # Sheet 1: MAIN_FIELDS
    ws1 = wb.create_sheet("MAIN_FIELDS")
    ws1.append(["Field","Value"])
    for k,v in main_fields.items():
        if v:  # Only include extracted data
            ws1.append([k,v])
    format_excel_sheet(ws1)

    # Sheet 2: TRANSACTIONS
    ws2 = wb.create_sheet("TRANSACTIONS")
    if transactions:
        ws2.append(list(transactions[0].keys()))
        for txn in transactions:
            ws2.append([v for v in txn.values()])
    format_excel_sheet(ws2)

    # Sheet 3: DAILY_BREAKDOWN
    ws3 = wb.create_sheet("DAILY_BREAKDOWN")
    if daily_breakdown:
        ws3.append(list(daily_breakdown[0].keys()))
        for row in daily_breakdown:
            ws3.append([v for v in row.values()])
    format_excel_sheet(ws3)

    # Sheet 4: WHERE_MONEY_WENT
    ws4 = wb.create_sheet("WHERE_MONEY_WENT")
    if destination_banks:
        ws4.append(list(destination_banks[0].keys()))
        for row in destination_banks:
            ws4.append([v for v in row.values()])
    format_excel_sheet(ws4)

    wb.save(output_excel)
