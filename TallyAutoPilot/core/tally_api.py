import requests
import pandas as pd
import yaml
import time
import xml.etree.ElementTree as ET
import re

CONFIG_FILE = "config.yaml"
EXCEL_FILE  = "data/sample_test_sales_with_company.xlsx"
REQUIRED    = ["CompanyName", "PartyName", "Amount", "Date"]

def load_config():
    """Load YAML config with keys: tally_host, tally_port."""
    with open(CONFIG_FILE) as f:
        return yaml.safe_load(f)

def send_xml(xml: str, host: str, port: int) -> str:
    """Send XML to Tally HTTP API and return response text."""
    resp = requests.post(f"http://{host}:{port}", data=xml, timeout=10)
    resp.raise_for_status()
    return resp.text

def create_ledger_from_party(row: dict, host: str, port: int):
    """
    Create a ledger based on PartyName under:
    - Sales Accounts if name == "Sales"
    - Sundry Debtors otherwise
    """
    name = row["PartyName"]
    company = row["CompanyName"]
    parent = "Sales Accounts" if name.strip().lower() == "sales" else "Sundry Debtors"
    print(f"  • Creating ledger '{name}' under '{parent}' for company '{company}'")

    xml = f"""
<ENVELOPE>
  <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>All Masters</REPORTNAME>
        <STATICVARIABLES>
          <SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>
        </STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>
        <TALLYMESSAGE xmlns:UDF="TallyUDF">
          <LEDGER NAME="{name}" ACTION="Create">
            <NAME.LIST><NAME>{name}</NAME></NAME.LIST>
            <PARENT>{parent}</PARENT>
            <DESCRIPTION>{name} ledger</DESCRIPTION>
            <CLOSINGBALANCE>0</CLOSINGBALANCE>
            <ISBILLWISEON>No</ISBILLWISEON>
          </LEDGER>
        </TALLYMESSAGE>
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>
"""
    return send_xml(xml, host, port)

def post_to_tally(row, config):
    """Import a single voucher, auto-creating only the party ledger if missing."""
    # Ensure dict
    if hasattr(row, "to_dict"):
        row = row.to_dict()

    # Validate
    for c in REQUIRED:
        if c not in row or pd.isna(row[c]):
            raise ValueError(f"Missing required column: {c}")

    host    = config["tally_host"]
    port    = config["tally_port"]
    company = row["CompanyName"]

    # Format date YYYYMMDD
    date = row["Date"]
    if hasattr(date, "strftime"):
        date = date.strftime("%Y%m%d")
    else:
        date = str(date).replace("-", "")

    party  = row["PartyName"]
    amount = row["Amount"]

    # Voucher XML (Sales ledger must already exist)
    xml_voucher = f"""
<ENVELOPE>
  <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>Vouchers</REPORTNAME>
        <STATICVARIABLES>
          <SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>
        </STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>
        <TALLYMESSAGE xmlns:UDF="TallyUDF">
          <VOUCHER VCHTYPE="Sales" ACTION="Create">
            <DATE>{date}</DATE>
            <NARRATION>Auto Entry</NARRATION>
            <VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>
            <PARTYLEDGERNAME>{party}</PARTYLEDGERNAME>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>{party}</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-{amount}</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Sales</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>{amount}</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
          </VOUCHER>
        </TALLYMESSAGE>
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>
"""

    # Send & handle missing‑ledger for the party only
    resp = send_xml(xml_voucher, host, port)
    root = ET.fromstring(resp)
    err  = root.find(".//LINEERROR")

    if err is None:
        return resp

    msg = err.text or ""
    # Check for missing date error
    if "Voucher date is missing" in msg:
        raise RuntimeError(f"Unexpected import error: {msg}")

    match = re.search(r"'([^']+)' does not exist", msg)
    if not match:
        raise RuntimeError(f"Unexpected import error: {msg}")

    missing = match.group(1)
    if missing == party:
        # Auto-create party ledger and retry once
        create_ledger_from_party(row, host, port)
        time.sleep(1)
        return send_xml(xml_voucher, host, port)
    else:
        # Any other missing ledger (e.g. "Sales") must be created manually
        raise RuntimeError(
            f"Tally reports missing ledger '{missing}'.\n"
            "Please create this ledger manually under the correct group and retry."
        )

def fetch_day_book(date, company, config):
    """Export and parse Day Book vouchers for a given date & company."""
    xml = f"""
<ENVELOPE>
  <HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>
  <BODY><EXPORTDATA>
    <REQUESTDESC>
      <REPORTNAME>Day Book</REPORTNAME>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVFROMDATE>{date}</SVFROMDATE>
        <SVTODATE>{date}</SVTODATE>
        <SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>
      </STATICVARIABLES>
    </REQUESTDESC>
  </EXPORTDATA></BODY>
</ENVELOPE>
"""
    resp = send_xml(xml, config["tally_host"], config["tally_port"])
    cleaned = re.sub(r"[^\x09\x0A\x0D\x20-\x7F]", "", resp)
    root = ET.fromstring(cleaned)
    return root.findall(".//VOUCHER")

if __name__ == "__main__":
    cfg = load_config()
    df  = pd.read_excel(EXCEL_FILE)

    print("=== IMPORTING VOUCHERS ===")
    for idx, row in df.iterrows():
        try:
            out = post_to_tally(row, cfg)
            # Only print output if not error
            print(f"Row {idx+1} OK\n")
        except Exception as e:
            print(f"Row {idx+1} ERROR: {e}\n")

    # Verify
    first    = df.iloc[0]
    date_str = pd.to_datetime(first["Date"]).strftime("%Y%m%d")
    company  = first["CompanyName"]
    time.sleep(2)

    print(f"=== VERIFYING DAY BOOK FOR {date_str} ===")
    vouchers = fetch_day_book(date_str, company, cfg)
    if not vouchers:
        print(f"No vouchers found for {date_str}")
    else:
        print(f"Found {len(vouchers)} voucher(s):")
        for v in vouchers:
            d = v.findtext("DATE", default="")
            ledger = v.find(".//LEDGERNAME")
            amt    = v.find(".//AMOUNT")
            print(f"  • {d} | {(ledger.text if ledger is not None else '')} | {(amt.text if amt is not None else '')}")
