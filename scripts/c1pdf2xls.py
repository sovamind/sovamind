import pandas as pd
import pdfplumber
from datetime import datetime

files = [
    "Statement_012025_0813.pdf",
    "Statement_022025_0813.pdf",
    "Statement_032025_0813.pdf",
    "Statement_042025_0813.pdf",
    "Statement_052025_0813.pdf"
]

rows = []

def parse_date(d):
    try:
        return datetime.strptime(d + " 2025", "%b %d %Y")
    except:
        return None

for file in files:
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            lines = text.split("\n")

            for line in lines:
                parts = line.split()

                # crude filter: must have 2 dates at start and a dollar amount at end
                if len(parts) < 4:
                    continue

                if parts[0].isalpha() and parts[2].isalpha():
                    try:
                        trans_date = parse_date(parts[0] + " " + parts[1])
                        post_date = parse_date(parts[2] + " " + parts[3])

                        amount_str = parts[-1].replace("$","").replace(",","")
                        amount = float(amount_str)

                        description = " ".join(parts[4:-1])

                        if trans_date:
                            rows.append({
                                "Transaction Date": trans_date,
                                "Post Date": post_date,
                                "Description": description,
                                "Amount": amount
                            })

                    except:
                        continue

df = pd.DataFrame(rows)
df = df.drop_duplicates()
df = df.sort_values("Transaction Date")

df.to_excel("FULL_credit_card_transactions.xlsx", index=False)

print("Done.")
