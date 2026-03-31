import pandas as pd
import json
from datetime import datetime, date

# ---- CONFIG ----
SOURCE_FILE = "NCD_MasterLog.xlsm"
SHEET_NAME = "MasterLog"
OUTPUT_JSON = "MasterLog.json"

TODAY = date.today()

# ---- HELPER FUNCTIONS ----
def parse_date(x):
    if pd.isna(x):
        return ""
    if isinstance(x, datetime):
        return x.date().isoformat()
    if isinstance(x, date):
        return x.isoformat()
    try:
        return pd.to_datetime(x).date().isoformat()
    except:
        return ""

def detect_type(ref):
    if not isinstance(ref, str):
        return ""
    ref_upper = ref.upper()
    for t in ["AQW", "AQO", "COR", "INV", "TOF"]:
        if t in ref_upper:
            return t
    return ""

# ---- LOAD DATA ----
df = pd.read_excel(SOURCE_FILE, sheet_name=SHEET_NAME, engine="openpyxl")

# Required columns:
# Ref. | Status | Date received | Response Due | Service Area | Subject | Lead | Approver | Closure Date | On Time?

# Rename columns internally for clean mapping
df2 = df.rename(columns={
    "Ref.": "ref",
    "Status": "status",
    "Date received": "received",
    "Response Due": "due",
    "Service Area": "area",
    "Subject": "subject",
    "Lead": "lead",
    "Approver": "approver",
    "Closure Date": "closed",
    "On Time?": "onTime"
})

records = []

for _, row in df2.iterrows():

    ref = row["ref"]
    status = row["status"]
    received = parse_date(row["received"])
    due = parse_date(row["due"])
    closed = parse_date(row["closed"])
    area = row["area"] if isinstance(row["area"], str) else ""
    lead = row["lead"] if isinstance(row["lead"], str) else ""
    approver = row["approver"] if isinstance(row["approver"], str) else ""
    subject = row["subject"] if isinstance(row["subject"], str) else ""
    onTime = bool(row["onTime"]) if not pd.isna(row["onTime"]) else False

    # Determine type
    type_val = detect_type(ref)

    # Determine open/closed/transferred
    status_upper = str(status).upper() if isinstance(status, str) else ""

    if "TRANSFER" in status_upper:
        isOpen = False
        isClosed = False
        isOverdue = False
        daysUntilDue = None
        turnaround = None
    else:
        # Compute open/closed
        isClosed = closed != ""
        isOpen = not isClosed

        # Compute daysUntilDue
        if due != "":
            due_date = datetime.fromisoformat(due).date()
            daysUntilDue = (due_date - TODAY).days
        else:
            daysUntilDue = None

        # Overdue logic
        isOverdue = isOpen and (daysUntilDue is not None and daysUntilDue < 0)

        # Turnaround (days between received and closed)
        if isClosed and received != "" and closed != "":
            rdate = datetime.fromisoformat(received).date()
            cdate = datetime.fromisoformat(closed).date()
            turnaround = (cdate - rdate).days
        else:
            turnaround = None

    record = {
        "ref": ref,
        "type": type_val,
        "status": status,
        "area": area,
        "lead": lead,
        "approver": approver,
        "subject": subject,
        "received": received,
        "due": due,
        "closed": closed,
        "onTime": onTime,
        "daysUntilDue": daysUntilDue,
        "isOpen": isOpen,
        "isClosed": isClosed,
        "isOverdue": isOverdue,
        "turnaround": turnaround
    }

    records.append(record)

# ---- WRITE JSON ----
with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(records, f, indent=2)

print(f"Export complete → {OUTPUT_JSON}")