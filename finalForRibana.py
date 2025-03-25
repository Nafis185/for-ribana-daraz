# STEP 1: Install required packages
!pip install pdfplumber pandas openpyxl

# STEP 2: Upload PDF
from google.colab import files
uploaded = files.upload()

# STEP 3: Import Libraries
import pdfplumber
import re
import pandas as pd

# Extract file path
pdf_path = list(uploaded.keys())[0]

# STEP 4: Extract Data
data = []

with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        if not text:
            continue
        lines = text.split('\n')

        order_id = ""
        order_date = ""
        deliver_to = ""
        phone = ""
        delivery_address = ""
        total = 0.0

        idx = 0
        while idx < len(lines):
            line = lines[idx]

            if "Order ID" in line:
                match_id = re.search(r"Order ID\s*(\d+)", line)
                if match_id:
                    order_id = match_id.group(1)
                elif idx + 1 < len(lines):
                    next_line_match = re.search(r"(\d+)", lines[idx + 1])
                    if next_line_match:
                        order_id = next_line_match.group(1)

            if "Order Date" in line:
                match_date = re.search(r"Order Date:\s*(.*)", line)
                if match_date:
                    order_date = match_date.group(1).strip()
                elif idx + 1 < len(lines):
                    order_date = lines[idx + 1].strip()

            if "Deliver To:" in line:
                dt_match = re.search(r"Deliver To:\s*(.*)", line)
                if dt_match:
                    dt_str = dt_match.group(1).strip()
                    phone_match = re.search(r"(?i)phone:\s*(\d+)", dt_str)
                    if phone_match:
                        phone = phone_match.group(1)
                        dt_str = re.sub(r"(?i)phone:\s*\d+", "", dt_str).strip()
                    deliver_to = dt_str

            if "Delivery Address:" in line:
                addr_lines = []

                # Line before "Delivery Address:" contains first part
                if idx - 1 >= 0:
                    prev_line = lines[idx - 1].strip()
                    if prev_line:
                        addr_lines.append(prev_line)

                # Line with "Delivery Address:"
                split_line = line.split("Delivery Address:")
                if len(split_line) > 1 and split_line[1].strip():
                    addr_lines.append(split_line[1].strip())

                # Collect following lines until "Bill To" or "Billing Address"
                j = idx + 1
                while j < len(lines):
                    next_line = lines[j].strip()
                    if re.search(r"Bill To|Billing Address|Phone:", next_line, re.IGNORECASE):
                        break
                    if next_line:
                        addr_lines.append(next_line)
                    j += 1

                delivery_address = ' '.join(addr_lines)

            if "Total:" in line:
                total_match = re.search(r'Total:\s*([\d,]+\.\d+|\d+)', line)
                if total_match:
                    total = float(total_match.group(1).replace(',', ''))

            idx += 1

        if order_id:
            data.append({
                "Order ID": order_id,
                "Order Date": order_date,
                "Deliver To": deliver_to,
                "Phone": phone,
                "Delivery Address": delivery_address,
                "Total": total
            })

# STEP 5: Create DataFrame
df = pd.DataFrame(data)

# STEP 6: Clean Illegal Characters
def clean_excel_string(text):
    if pd.isna(text):
        return text
    return re.sub(r'[\x00-\x1F\x7F-\x9F]', '', str(text))

for col in df.columns:
    df[col] = df[col].apply(clean_excel_string)

# STEP 7: Save & Download Excel
output_excel = "Extracted_Order_Details_Updated.xlsx"
df.to_excel(output_excel, index=False)

files.download(output_excel)
