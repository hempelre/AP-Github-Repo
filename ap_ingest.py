"""
ap_ingest.py
------------
Daily AP Vendor Aging ingest.

Two ways to run:

  Scheduled (GitHub Actions): no CLI args. Pulls the latest attachment from
  the "AP Vendor Aging" folder in the data@hempelcompanies.com shared mailbox
  via Microsoft Graph, parses it, reconciles it, and (if DO_INGEST=true)
  writes to dbo.AP_Invoices / AP_LineItems / AP_Totals. On success, moves the
  source email to the "Old AP Vendor Aging" folder (if DO_MOVE=true).

  Manual / local testing: pass a file path on the command line to skip Graph
  entirely and parse a local file, same as before.

      python ap_ingest.py "Vendor_Aging_Report (5).xls"           # parse only
      python ap_ingest.py "Vendor_Aging_Report (5).xls" --commit  # write to DB

Env vars (see .env.template):
  AUTH_METHOD           "client_secret" (GitHub Actions has no managed identity)
  GRAPH_TENANT_ID / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET
  SHARED_MAILBOX        e.g. data@hempelcompanies.com
  SOURCE_FOLDER         e.g. "AP Vendor Aging"
  DEST_FOLDER           e.g. "Old AP Vendor Aging"
  SQL_CONN_STR          required for DO_INGEST=true
  DO_INGEST             "true" to write to the DB (default: false -> preview only)
  DO_MOVE               "true" to move the email after a successful write
  TEST_MODE             "true" logs extra detail, does not change write behavior
  NO_RECONCILE_GUARD    "true" to write even if reconciliation shows a mismatch

One-time setup before the first write: run create_ap_invoices.sql so
dbo.AP_Invoices / AP_LineItems / AP_Totals exist.
"""

import os
import sys
import base64
import argparse
from io import BytesIO

import requests
import pandas as pd

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

SQL_CONN_STR = os.getenv("SQL_CONN_STR", "").strip().strip('"')

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _env_bool(name: str, default: bool = False) -> bool:
    val = os.getenv(name)
    if val is None:
        return default
    return val.strip().lower() in ("1", "true", "yes", "on")


# ============================================================
# GRAPH FETCH  (client_secret auth — no managed identity in GH Actions)
# ============================================================
def get_access_token() -> str:
    from azure.identity import ClientSecretCredential

    tenant_id = os.environ["GRAPH_TENANT_ID"]
    client_id = os.environ["GRAPH_CLIENT_ID"]
    client_secret = os.environ["GRAPH_CLIENT_SECRET"]

    cred = ClientSecretCredential(tenant_id, client_id, client_secret)
    token = cred.get_token("https://graph.microsoft.com/.default")
    return token.token


def _graph_get(token: str, url: str, params: dict = None) -> dict:
    resp = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        params=params or {},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def find_folder_id(token: str, mailbox: str, folder_name: str) -> str:
    """Look for folder_name at the top level of the mailbox, then one level
    under Inbox, since that's where these tend to get filtered to."""
    top = _graph_get(
        token,
        f"{GRAPH_BASE}/users/{mailbox}/mailFolders",
        params={"$filter": f"displayName eq '{folder_name}'", "$top": 1},
    )
    if top.get("value"):
        return top["value"][0]["id"]

    inbox = _graph_get(token, f"{GRAPH_BASE}/users/{mailbox}/mailFolders/inbox")
    children = _graph_get(
        token,
        f"{GRAPH_BASE}/users/{mailbox}/mailFolders/{inbox['id']}/childFolders",
        params={"$filter": f"displayName eq '{folder_name}'", "$top": 1},
    )
    if children.get("value"):
        return children["value"][0]["id"]

    raise ValueError(f"Could not find mail folder '{folder_name}' in {mailbox}.")


def get_latest_message_with_attachment(token: str, mailbox: str, folder_id: str) -> dict:
    result = _graph_get(
        token,
        f"{GRAPH_BASE}/users/{mailbox}/mailFolders/{folder_id}/messages",
        params={
            "$orderby": "receivedDateTime desc",
            "$top": 10,
            "$select": "id,subject,receivedDateTime,hasAttachments",
        },
    )
    for msg in result.get("value", []):
        if msg.get("hasAttachments"):
            return msg
    raise ValueError("No recent message with an attachment found in the source folder.")

def get_first_attachment_bytes(token: str, mailbox: str, message_id: str) -> bytes:
    result = _graph_get(
        token, f"{GRAPH_BASE}/users/{mailbox}/messages/{message_id}/attachments"
    )
    for att in result.get("value", []):
        if att.get("contentBytes"):
            return base64.b64decode(att["contentBytes"])
    raise ValueError("Message had hasAttachments=True but no downloadable file attachment.")


def move_message(token: str, mailbox: str, message_id: str, dest_folder_id: str) -> None:
    resp = requests.post(
        f"{GRAPH_BASE}/users/{mailbox}/messages/{message_id}/move",
        headers={"Authorization": f"Bearer {token}"},
        json={"destinationId": dest_folder_id},
        timeout=30,
    )
    resp.raise_for_status()


def fetch_latest_report_from_mailbox():
    """Returns (file_bytes, message_id, mailbox, dest_folder_id_or_None)."""
    mailbox = os.environ["SHARED_MAILBOX"]
    source_folder = os.environ.get("SOURCE_FOLDER", "AP Vendor Aging")
    dest_folder = os.environ.get("DEST_FOLDER", "Old AP Vendor Aging")
    do_move = _env_bool("DO_MOVE", False)

    token = get_access_token()
    source_id = find_folder_id(token, mailbox, source_folder)
    msg = get_latest_message_with_attachment(token, mailbox, source_id)
    print(f"Found message: '{msg.get('subject')}' ({msg['id']}) received {msg.get('receivedDateTime')}")

    file_bytes = get_first_attachment_bytes(token, mailbox, msg["id"])

    dest_id = None
    if do_move:
        dest_id = find_folder_id(token, mailbox, dest_folder)

    return file_bytes, msg["id"], mailbox, token, dest_id


# ============================================================
# PARSER  (name-based; handles Detail 15-col and Summary 9-col)
# ============================================================
CANON = {
    "property":     ["property", "property only"],
    "vendor_id":    ["vendor id"],
    "vendor_name":  ["vendor name"],
    "ap_invoice":   ["ap invoice"],
    "on_hold":      ["ap invoices on hold"],
    "gl_post_date": ["gl posting date"],
    "inv_date":     ["ap invoice date"],
    "due_date":     ["due date"],
    "days_aged":    ["days aged"],
    "b_0_30":       ["0-30"],
    "b_31_60":      ["31-60"],
    "b_61_90":      ["61-90"],
    "b_91_119":     ["91-119"],
    "b_120":        ["120-", "120+", "120"],
    "total":        ["total"],
}
BUCKETS = ["b_0_30", "b_31_60", "b_61_90", "b_91_119", "b_120", "total"]
BUCKET_LABELS = {
    "b_0_30": "0-30", "b_31_60": "31-60", "b_61_90": "61-90",
    "b_91_119": "91-119", "b_120": "120+", "total": "Total",
}


def clean_money(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.replace({"": None, "nan": None, "None": None, "-": None})
    s = (s.str.replace(r"[\$,]", "", regex=True)
           .str.replace(r"^\((.*)\)$", r"-\1", regex=True))
    s = s.apply(lambda x: "-" + x[:-1] if isinstance(x, str) and x.endswith("-") else x)
    return pd.to_numeric(s, errors="coerce")


def clean_text(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    return s.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})


def clean_date(series: pd.Series):
    return pd.to_datetime(series, errors="coerce").dt.date


def resolve_columns(df: pd.DataFrame) -> dict:
    def last_level(col):
        return col[-1] if isinstance(col, tuple) else col
    normalized = {last_level(c).strip().lower(): c for c in df.columns}
    resolved = {}
    for key, aliases in CANON.items():
        for alias in aliases:
            if alias in normalized:
                resolved[key] = normalized[alias]
                break
    return resolved


def extract_as_of_date(header_df: pd.DataFrame):
    for _, row in header_df.iterrows():
        row_str = row.astype(str)
        mask = row_str.str.contains("As of Date", case=False, na=False)
        if mask.any():
            col = mask.idxmax()
            for val in row.loc[col + 1:].dropna().astype(str):
                try:
                    return pd.to_datetime(val).date()
                except Exception:
                    continue
    raise ValueError("Could not find a valid 'As of Date' in the report header.")


def parse_report(file_bytes: bytes):
    tables = pd.read_html(BytesIO(file_bytes))
    if len(tables) < 2:
        raise ValueError(f"Expected >=2 HTML tables, found {len(tables)}.")

    header_df = tables[0]
    df = tables[1].copy()

    as_of_date = extract_as_of_date(header_df)
    cols = resolve_columns(df)

    required = ["property", "vendor_id", "vendor_name", *BUCKETS]
    missing = [k for k in required if k not in cols]
    if missing:
        raise ValueError(
            "Report format not recognized. Missing expected columns: "
            f"{missing}. Parsed header was: {list(df.columns)}"
        )

    is_detail = "ap_invoice" in cols
    fmt = "DETAIL" if is_detail else "SUMMARY"

    prop_raw = clean_text(df[cols["property"]])
    is_subtotal = prop_raw.fillna("").str.lower().str.startswith("total for ")
    is_grand = df.apply(
        lambda r: r.astype(str).str.strip().str.lower().eq("grand totals").any(),
        axis=1,
    )

    prop_ffill = prop_raw.where(~(is_subtotal | is_grand)).ffill()
    vid_ffill = clean_text(df[cols["vendor_id"]]).ffill()
    vname_ffill = clean_text(df[cols["vendor_name"]]).ffill()

    for key in BUCKETS:
        df[key] = clean_money(df[cols[key]])

    invoices = pd.DataFrame()
    if is_detail:
        inv_no = clean_text(df[cols["ap_invoice"]])
        is_line = (~is_subtotal) & (~is_grand) & df["total"].notna()
        invoices = pd.DataFrame({
            "AsOfDate":      as_of_date,
            "PropertyName":  prop_ffill,
            "VendorNumber":  vid_ffill,
            "VendorName":    vname_ffill,
            "APInvoice":     inv_no,
            "OnHold":        clean_text(df[cols["on_hold"]]) if "on_hold" in cols else pd.NA,
            "GLPostingDate": clean_date(df[cols["gl_post_date"]]) if "gl_post_date" in cols else pd.NaT,
            "APInvoiceDate": clean_date(df[cols["inv_date"]]) if "inv_date" in cols else pd.NaT,
            "DueDate":       clean_date(df[cols["due_date"]]) if "due_date" in cols else pd.NaT,
            "DaysAged":      pd.to_numeric(df[cols["days_aged"]], errors="coerce") if "days_aged" in cols else pd.NA,
            **{BUCKET_LABELS[k]: df[k] for k in BUCKETS},
        })[is_line].reset_index(drop=True)

    if is_detail:
        base = invoices.copy()
    else:
        vend_mask = (~is_subtotal) & (~is_grand) & vname_ffill.notna() & df["total"].notna()
        base = pd.DataFrame({
            "PropertyName": prop_ffill,
            "VendorNumber": vid_ffill,
            "VendorName":   vname_ffill,
            **{BUCKET_LABELS[k]: df[k] for k in BUCKETS},
        })[vend_mask].reset_index(drop=True)
        base.insert(0, "AsOfDate", as_of_date)

    grp_cols = ["AsOfDate", "PropertyName", "VendorNumber", "VendorName"]
    sum_cols = [BUCKET_LABELS[k] for k in BUCKETS]
    vendor_rollup = (base.groupby(grp_cols, as_index=False)[sum_cols].sum()
                     if not base.empty else base)

    stated = pd.DataFrame()
    if is_subtotal.any():
        stated = pd.DataFrame({
            "PropertyName": prop_raw[is_subtotal].str.replace(
                r"(?i)^total for\s*", "", regex=True).str.strip(),
            **{BUCKET_LABELS[k]: df[k][is_subtotal] for k in BUCKETS},
        }).drop_duplicates(subset=["PropertyName"], keep="last").reset_index(drop=True)

    computed = (vendor_rollup.groupby("PropertyName", as_index=False)[sum_cols].sum()
                if not vendor_rollup.empty else pd.DataFrame(columns=["PropertyName", *sum_cols]))
    merged = computed.merge(stated, on="PropertyName", how="outer",
                            suffixes=("_computed", "_stated"))
    property_totals = pd.DataFrame({"AsOfDate": as_of_date,
                                    "PropertyName": merged["PropertyName"]})
    for lbl in sum_cols:
        sc, cc = f"{lbl}_stated", f"{lbl}_computed"
        if sc in merged and cc in merged:
            property_totals[lbl] = merged[sc].fillna(merged[cc])
        elif sc in merged:
            property_totals[lbl] = merged[sc]
        else:
            property_totals[lbl] = merged[cc]
    property_totals = (property_totals[property_totals["PropertyName"].notna()]
                       .sort_values("PropertyName")
                       .drop_duplicates(subset=["AsOfDate", "PropertyName"], keep="last")
                       .reset_index(drop=True))

    meta = {"format": fmt, "as_of_date": as_of_date, "raw_rows": len(df),
            "stated": stated, "is_grand": is_grand, "df": df, "sum_cols": sum_cols}
    return invoices, vendor_rollup, property_totals, meta


def reconcile(invoices, vendor_rollup, property_totals, meta, tol=0.01):
    print("\n" + "=" * 72 + "\nRECONCILIATION\n" + "=" * 72)
    sum_cols = meta["sum_cols"]
    stated = meta["stated"]
    ok = True
    if not stated.empty and not invoices.empty:
        comp = invoices.groupby("PropertyName", as_index=False)[sum_cols].sum()
        chk = stated.merge(comp, on="PropertyName", how="outer",
                           suffixes=("_stated", "_computed"))
        chk["Total_diff"] = chk["Total_stated"].fillna(0) - chk["Total_computed"].fillna(0)
        bad = chk[chk["Total_diff"].abs() > tol]
        print(f"Properties reconciled against 'Total for' rows: {len(chk)}")
        print(f"  within tolerance (<= {tol}): {len(chk) - len(bad)}")
        print(f"  MISMATCHES: {len(bad)}")
        if len(bad):
            ok = False
            print(bad[["PropertyName", "Total_stated", "Total_computed", "Total_diff"]].to_string(index=False))
    else:
        print("No stated subtotal rows and/or no invoice detail to reconcile.")

    df, is_grand = meta["df"], meta["is_grand"]
    if is_grand.any():
        grand = df[is_grand].iloc[0]
        total_col = [c for c in df.columns
                     if (c[-1] if isinstance(c, tuple) else c).strip().lower() == "total"][0]
        grand_total = clean_money(pd.Series([grand[total_col]])).iloc[0]
        detail_total = invoices["Total"].sum() if not invoices.empty else vendor_rollup["Total"].sum()
        diff = (grand_total or 0) - (detail_total or 0)
        flag = "OK" if abs(diff) <= tol else "*** MISMATCH ***"
        if abs(diff) > tol:
            ok = False
        print(f"\nGrand total row:      {grand_total:,.2f}")
        print(f"Sum of parsed detail: {detail_total:,.2f}")
        print(f"Difference:           {diff:,.2f}  {flag}")
    return ok


# ============================================================
# DB LOAD
# ============================================================
INVOICE_COLS = ["AsOfDate", "PropertyName", "VendorNumber", "VendorName", "APInvoice", "OnHold",
                "GLPostingDate", "APInvoiceDate", "DueDate", "DaysAged",
                "0-30", "31-60", "61-90", "91-119", "120+", "Total"]
LINEITEM_COLS = ["AsOfDate", "PropertyName", "VendorNumber", "VendorName",
                 "0-30", "31-60", "61-90", "91-119", "120+", "Total"]
TOTAL_COLS = ["AsOfDate", "PropertyName", "0-30", "31-60", "61-90", "91-119", "120+", "Total"]

INSERT_INVOICES = """
INSERT INTO dbo.AP_Invoices
 (AsOfDate, PropertyName, VendorNumber, VendorName, APInvoice, OnHold,
  GLPostingDate, APInvoiceDate, DueDate, DaysAged,
  [0-30], [31-60], [61-90], [91-119], [120+], Total)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
"""
INSERT_LINEITEMS = """
INSERT INTO dbo.AP_LineItems
 (AsOfDate, PropertyName, VendorNumber, VendorName,
  [0-30], [31-60], [61-90], [91-119], [120+], Total)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
"""
INSERT_TOTALS = """
INSERT INTO dbo.AP_Totals
 (AsOfDate, PropertyName, [0-30], [31-60], [61-90], [91-119], [120+], Total)
VALUES (?, ?, ?, ?, ?, ?, ?, ?)
"""


def get_sql_connection():
    if not SQL_CONN_STR:
        raise ValueError("SQL_CONN_STR is not set (required to write to the DB).")
    import pyodbc
    return pyodbc.connect(SQL_CONN_STR)


def _rows(df: pd.DataFrame, cols: list) -> list:
    if df.empty:
        return []
    sub = df[cols].copy()
    if "DaysAged" in sub.columns:
        sub["DaysAged"] = pd.to_numeric(sub["DaysAged"], errors="coerce").astype("Int64")
    sub = sub.astype(object).where(pd.notnull(sub), None)
    return [tuple(r) for r in sub.to_numpy()]


def write_snapshot(as_of_date, invoices, vendor_rollup, property_totals, replace=True):
    conn = get_sql_connection()
    cur = conn.cursor()
    cur.fast_executemany = True
    try:
        if replace:
            cur.execute("DELETE FROM dbo.AP_Invoices  WHERE AsOfDate = ?", as_of_date)
            cur.execute("DELETE FROM dbo.AP_LineItems WHERE AsOfDate = ?", as_of_date)
            cur.execute("DELETE FROM dbo.AP_Totals    WHERE AsOfDate = ?", as_of_date)

        inv_rows = _rows(invoices, INVOICE_COLS)
        if inv_rows:
            cur.executemany(INSERT_INVOICES, inv_rows)
        cur.executemany(INSERT_LINEITEMS, _rows(vendor_rollup, LINEITEM_COLS))
        cur.executemany(INSERT_TOTALS,    _rows(property_totals, TOTAL_COLS))

        conn.commit()
        print(f"\nCOMMITTED {as_of_date}: {len(inv_rows)} invoices, "
              f"{len(vendor_rollup)} vendor rows, {len(property_totals)} property totals.")
    except Exception:
        conn.rollback()
        print("\nLoad FAILED — transaction rolled back, nothing written.")
        raise
    finally:
        cur.close()
        conn.close()


# ============================================================
# MAIN
# ============================================================
def process_and_report(file_bytes: bytes, do_write: bool, no_reconcile_guard: bool, dump_csv: bool):
    invoices, vendor_rollup, property_totals, meta = parse_report(file_bytes)

    print("=" * 72)
    print(f"FORMAT:  {meta['format']}   AS OF: {meta['as_of_date']}   RAW ROWS: {meta['raw_rows']}")
    print(f"invoices={len(invoices)}  vendor_rollup={len(vendor_rollup)}  property_totals={len(property_totals)}")

    if not invoices.empty:
        print("\n--- invoices (first 15) ---")
        show = ["PropertyName", "VendorNumber", "VendorName", "APInvoice",
                "GLPostingDate", "DueDate", "DaysAged", "120+", "Total"]
        print(invoices[show].head(15).to_string(index=False))
    print("\n--- vendor_rollup (first 15) ---")
    print(vendor_rollup.head(15).to_string(index=False))
    print("\n--- property_totals (first 15) ---")
    print(property_totals.head(15).to_string(index=False))

    recon_ok = reconcile(invoices, vendor_rollup, property_totals, meta)

    if dump_csv:
        if not invoices.empty:
            invoices.to_csv("parsed_invoices.csv", index=False)
        vendor_rollup.to_csv("parsed_vendor_rollup.csv", index=False)
        property_totals.to_csv("parsed_property_totals.csv", index=False)
        print("\nCSVs written to current directory.")

    if not do_write:
        print("\nParse only — nothing written.")
        return True

    if not recon_ok and not no_reconcile_guard:
        print("\nABORTING WRITE: reconciliation failed. "
              "Set NO_RECONCILE_GUARD=true to force, or investigate the file.")
        return False

    write_snapshot(meta["as_of_date"], invoices, vendor_rollup, property_totals, replace=True)
    return True


def main():
    ap = argparse.ArgumentParser(description="Parse a RealPage Vendor Aging file.")
    ap.add_argument("path", nargs="?", default=None,
                    help="Local file path (manual/testing mode). Omit to fetch from the mailbox.")
    ap.add_argument("--commit", action="store_true",
                    help="Write results to the database (local mode only; scheduled mode uses DO_INGEST).")
    ap.add_argument("--csv", action="store_true",
                    help="Dump the parsed frames to CSV in the current directory.")
    ap.add_argument("--no-reconcile-guard", action="store_true")
    args = ap.parse_args()

    if args.path:
        # ---- local/manual mode: unchanged from before ----
        with open(args.path, "rb") as f:
            file_bytes = f.read()
        print(f"FILE: {args.path}")
        ok = process_and_report(file_bytes, args.commit, args.no_reconcile_guard, args.csv)
        if not ok:
            raise SystemExit(1)
        return

    # ---- scheduled mode: fetch from the mailbox via Graph ----
    do_ingest = _env_bool("DO_INGEST", False)
    do_move = _env_bool("DO_MOVE", False)
    no_guard = _env_bool("NO_RECONCILE_GUARD", False)
    test_mode = _env_bool("TEST_MODE", False)

    print(f"AUTH_METHOD=client_secret  DO_INGEST={do_ingest}  DO_MOVE={do_move}  TEST_MODE={test_mode}")

    file_bytes, message_id, mailbox, token, dest_folder_id = fetch_latest_report_from_mailbox()

    ok = process_and_report(file_bytes, do_write=do_ingest, no_reconcile_guard=no_guard, dump_csv=False)

    if not ok:
        raise SystemExit(1)

    if do_ingest and do_move and dest_folder_id:
        move_message(token, mailbox, message_id, dest_folder_id)
        print(f"Moved source email to destination folder.")
    elif do_move and not do_ingest:
        print("DO_MOVE is true but DO_INGEST is false — leaving email in place so it's reprocessed next run.")


if __name__ == "__main__":
    main()
