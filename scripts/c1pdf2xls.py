import pandas as pd
import pdfplumber
import re
import argparse
import glob
from datetime import datetime

# --- CLI ARGUMENTS ---
parser = argparse.ArgumentParser(description="Extract transactions from credit card PDFs")

parser.add_argument("-i", "--input", nargs="+", help="Input PDF files (space separated)")
parser.add_argument("-o", "--output", help="Output Excel file")

args = parser.parse_args()

# --- INPUT FILES ---
if args.input:
    files = args.input
else:
    print("No input filenames supplied, looking for all PDFs in same folder")
    files = glob.glob("*.pdf")

# --- OUTPUT FILE ---
if args.output:
    output_file = args.output
else:
    print("No output filename supplied, using transactions.xlsx")
    output_file = "transactions.xlsx"

rows = []

# --- EXTRACT YEAR RANGE FROM HEADER ---
def extract_years(text):
    match = re.search(
        r"([A-Za-z]{3} \d{1,2}, (\d{4}))\s*-\s*([A-Za-z]{3} \d{1,2}, (\d{4}))",
        text
    )
    if match:
        return int(match.group(2)), int(match.group(4))
    return None, None

# --- PARSE DATE USING STATEMENT CONTEXT ---
def parse_date(month_day, start_year, end_year):
    month = datetime.strptime(month_day, "%b %d").month

    # If spanning year boundary (Dec → Jan)
    if start_year != end_year:
        year = start_year if month == 12 else end_year
    else:
        year = start_year

    return datetime.strptime(f"{month_day} {year}", "%b %d %Y")

# --- PROCESS FILES ---
for file in files:
    print(f"Processing: {file}")

    with pdfplumber.open(file) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)

        start_year, end_year = extract_years(full_text)

        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            lines = text.split("\n")

            for line in lines:
                parts = line.split()

                # Must look like: Mar 17 Mar 18 ... $123.45
                if len(parts) < 6:
                    continue

                if not parts[0].isalpha():
                    continue

                try:
                    trans_date_raw = parts[0] + " " + parts[1]
                    post_date_raw = parts[2] + " " + parts[3]

                    amount_str = parts[-1].replace("$", "").replace(",", "")
                    amount = float(amount_str)

                    description = " ".join(parts[4:-1])

                    trans_date = parse_date(trans_date_raw, start_year, end_year)
                    post_date = parse_date(post_date_raw, start_year, end_year)

                    rows.append({
                        "Transaction Date": trans_date,
                        "Post Date": post_date,
                        "Description": description,
                        "Amount": amount
                    })

                except:
                    continue

# --- FINAL DATAFRAME ---
df = pd.DataFrame(rows)

df = df.drop_duplicates()
df = df.sort_values("Transaction Date")

df.to_excel(output_file, index=False)

print(f"Done. Wrote {len(df)} rows to {output_file}")
