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
    with open(CONFIG_FILE) as f:
        return yaml.safe_load(f)

def send_xml(xml: str, host: str, port: int) -> str:
    resp = requests.post(f"http://{host}:{port}", data=xml, timeout=10)
    resp.raise_for_status()
    return resp.text

def create_party_ledger(name: str, company: str, host: str, port: int):
    """Create a missing party ledger under Sundry Debtors."""
    print(f"  • Creating party ledger '{name}' under 'Sundry Debtors' for company '{company}'")
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
            <PARENT>Sundry Debtors</PARENT>
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

def post_to_tally(row: dict, config: dict) -> str:
    """Post a voucher, auto-creating only missing party ledger."""
    # Validate
    for col in REQUIRED:
        if col not in row or pd.isna(row[col]):
            raise ValueError(f"Missing required column: {col}")

    host    = config["tally_host"]
    port    = config["tally_port"]
    company = row["CompanyName"]

    # Format date as YYYYMMDD
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

    # Send & handle missing party ledger
    resp = send_xml(xml_voucher, host, port)
    root = ET.fromstring(resp)
    err  = root.find(".//LINEERROR")

    if err is None:
        return resp

    msg = err.text or ""
    match = re.search(r"'([^']+)' does not exist", msg)
    if not match:
        raise RuntimeError(f"Unexpected import error: {msg}")

    missing = match.group(1)
    if missing == party:
        # Auto-create party ledger and retry once
        create_party_ledger(party, company, host, port)
        time.sleep(1)
        return send_xml(xml_voucher, host, port)
    else:
        # Any other missing ledger must be created manually
        raise RuntimeError(
            f"Tally reports missing ledger '{missing}'.\n"
            "Please create this ledger manually under the correct group and retry."
        )

def fetch_day_book(date: str, company: str, config: dict):
    """Fetch Day Book vouchers for the given date & company, handling XML parse errors gracefully. Returns a list of dicts with DATE, LEDGERNAME, and AMOUNT."""
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
    try:
        resp = send_xml(xml, config["tally_host"], config["tally_port"])
    except Exception as e:
        print(f"ERROR: Failed to fetch from Tally: {e}")
        return []
    cleaned = re.sub(r"[^\x09\x0A\x0D\x20-\x7F]", "", resp)
    try:
        root = ET.fromstring(cleaned)
    except ET.ParseError as e:
        print("ERROR: Failed to parse Tally XML response. This may be due to invalid characters in the response.")
        print(f"ParseError: {e}")
        print("Raw (cleaned) response snippet:")
        print(cleaned[:500] + ("..." if len(cleaned) > 500 else ""))
        return []
    vouchers = root.findall(".//VOUCHER")
    results = []
    for v in vouchers:
        d = v.findtext("DATE", default="")
        ledger = v.find(".//LEDGERNAME")
        amt = v.find(".//AMOUNT")
        results.append({
            "DATE": d,
            "LEDGERNAME": ledger.text if ledger is not None else "",
            "AMOUNT": amt.text if amt is not None else ""
        })
    return results

if __name__ == "__main__":
    cfg = load_config()
    df  = pd.read_excel(EXCEL_FILE)

    print("=== IMPORTING VOUCHERS ===")
    for idx, row in df.iterrows():
        row_data = row.to_dict()
        print(f"Row {idx+1}:")
        try:
            out = post_to_tally(row_data, cfg)
            print(out, "\n")
        except Exception as e:
            print("ERROR:", e, "\n")

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
            d = v.get("DATE", "")
            ledger = v.get("LEDGERNAME", "")
            amt    = v.get("AMOUNT", "")
            print(f"  • {d} | {ledger} | {amt}")
