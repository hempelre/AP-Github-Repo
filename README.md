# AP Ingest Pipeline

Daily ingest of the RealPage AP Vendor Aging report (email attachment) into
`dbo.AP_Invoices` / `AP_LineItems` / `AP_Totals` on Azure SQL.

Runs weekdays at 8:00 AM Central via GitHub Actions. Same architecture as the
Arpari and lease-expiration pipelines: Microsoft Graph (client_secret auth) →
pandas parse/reconcile → pyodbc write, with a mail-move-on-success step.

## One-time setup

1. **Mail folder.** In the `data@hempelcompanies.com` mailbox, confirm an
   `AP Vendor Aging` folder exists (top level or a child of Inbox) where the
   RealPage report lands, and create `Old AP Vendor Aging` as the archive
   destination.
2. **App registration.** This reuses the same Entra app registration as the
   Arpari/lease pipelines. Confirm it has `Mail.ReadWrite` app permission
   (admin consent granted) on the `data@hempelcompanies.com` mailbox.
3. **Database tables.** Confirm `dbo.AP_Invoices`, `dbo.AP_LineItems`, and
   `dbo.AP_Totals` already exist (per the original script's
   `create_ap_invoices.sql`, if you still have it).
4. **GitHub secrets** (repo Settings → Secrets and variables → Actions):
   - `GRAPH_TENANT_ID`
   - `GRAPH_CLIENT_ID`
   - `GRAPH_CLIENT_SECRET`
   - `SHARED_MAILBOX` → `data@hempelcompanies.com`
   - `SQL_CONN_STR`

## Testing before turning it loose

1. Copy `.env.template` to `.env`, fill in real values, leave
   `DO_INGEST=false` and `DO_MOVE=false`.
2. Run `python ap_ingest.py` locally — it will fetch the latest email in the
   source folder, parse it, print the reconciliation, and stop (no write, no
   move).
3. Once the parse/reconciliation looks right, flip `DO_INGEST=true` and
   re-run to confirm the DB write. Then flip `DO_MOVE=true` last, after
   you're confident it's finding the right email each time.
4. Push to GitHub, add the secrets above, and use **Run workflow** in the
   Actions tab to do one manual end-to-end run before trusting the cron.

## Notes

- The reconciliation guard aborts the write if the parsed totals don't match
  the report's own "Total for" / "Grand Totals" rows. Override with
  `NO_RECONCILE_GUARD=true` only if you've confirmed the mismatch is benign.
- `write_snapshot` deletes-then-inserts by `AsOfDate`, so a rerun on the same
  day's file is safe and won't duplicate rows.
- If `DO_MOVE=true` but `DO_INGEST=false`, the script deliberately leaves the
  email in place so it gets reprocessed on the next run.
- You can still run this manually against a local file for testing/backfill:
  `python ap_ingest.py "Vendor_Aging_Report (5).xls" --commit`
- Local Python needs `pandas`, `lxml` (for `pd.read_html`), `requests`,
  `azure-identity`, `pyodbc`, `python-dotenv` — see `requirements.txt`.
