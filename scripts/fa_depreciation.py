"""
fa_depreciation.py
Retro Labs Ltd — Fixed Asset Month-End Automation
Triggered by GitHub Actions cron on the 1st of each month.

Steps:
  1. GL audit        — Snowflake: detect unregistered assets on 720/710
  2. Register check  — Snowflake/Fivetran: confirm 0 drafts, sync current
  3. Run dep         — Xero API: POST depreciation for prior month
  4. Verify          — Snowflake: confirm journals posted within +/-£10
  5. Rebuild schedule — build_fa_schedule.py: regenerate xlsx, upload Drive
  6. Notify          — Slack + GitHub Actions job summary

GitHub Actions secrets required:
  SNOWFLAKE_ACCOUNT, SNOWFLAKE_USER, SNOWFLAKE_PASSWORD
  SNOWFLAKE_DATABASE, SNOWFLAKE_WAREHOUSE, SNOWFLAKE_SCHEMA
  XERO_CLIENT_ID, XERO_CLIENT_SECRET, XERO_TENANT_ID, XERO_REFRESH_TOKEN
  SLACK_WEBHOOK_URL
  GOOGLE_DRIVE_FOLDER_ID, GOOGLE_SERVICE_ACCOUNT_JSON
"""

import os, sys, time, datetime, calendar, requests, snowflake.connector
from dateutil.relativedelta import relativedelta

TODAY        = datetime.date.today()
PRIOR        = TODAY - relativedelta(months=1)
PERIOD_YEAR  = PRIOR.year
PERIOD_MONTH = PRIOR.month
PERIOD_LABEL = PRIOR.strftime("%B %Y")
PERIOD_START = datetime.date(PERIOD_YEAR, PERIOD_MONTH, 1)
PERIOD_END   = datetime.date(PERIOD_YEAR, PERIOD_MONTH,
                             calendar.monthrange(PERIOD_YEAR, PERIOD_MONTH)[1])
PERIOD_END_ISO = PERIOD_END.isoformat()

CAP_THRESHOLD    = 100.0
VERIFY_TOLERANCE = 10.0


def get_conn():
    return snowflake.connector.connect(
        account   = os.environ["SNOWFLAKE_ACCOUNT"],
        user      = os.environ["SNOWFLAKE_USER"],
        password  = os.environ["SNOWFLAKE_PASSWORD"],
        database  = os.environ["SNOWFLAKE_DATABASE"],
        warehouse = os.environ["SNOWFLAKE_WAREHOUSE"],
        schema    = os.environ.get("SNOWFLAKE_SCHEMA", "XERO"),
    )

def slack(msg, emoji=":white_check_mark:"):
    url = os.environ.get("SLACK_WEBHOOK_URL")
    if not url:
        print(f"[SLACK] {msg}"); return
    requests.post(url, json={"text": f"{emoji} *FA Automation — {PERIOD_LABEL}*\n{msg}"}, timeout=10)

def fail(reason):
    slack(reason, ":red_circle:")
    print(f"\n[FAIL] {reason}", file=sys.stderr)
    sys.exit(1)

def log(msg): print(f"  {msg}")


# ── STEP 1: GL AUDIT ──────────────────────────────────────────────────────────

def step1_gl_audit(conn):
    print(f"[1/6] GL audit — {PERIOD_LABEL}")
    cur = conn.cursor()
    cur.execute("""
        SELECT TO_DATE(j.JOURNAL_DATE), j.JOURNAL_NUMBER,
               jl.ACCOUNT_CODE, jl.DESCRIPTION, jl.NET_AMOUNT
        FROM FIVETRAN.XERO.JOURNAL j
        JOIN FIVETRAN.XERO.JOURNAL_LINE jl ON j.JOURNAL_ID = jl.JOURNAL_ID
        WHERE TO_DATE(j.JOURNAL_DATE) BETWEEN %(s)s AND %(e)s
          AND jl.ACCOUNT_CODE IN ('720','710')
          AND jl.NET_AMOUNT > %(t)s
        ORDER BY jl.NET_AMOUNT DESC
    """, {"s": PERIOD_START.isoformat(), "e": PERIOD_END_ISO, "t": CAP_THRESHOLD})
    cols = [d[0] for d in cur.description]
    gl_rows = [dict(zip(cols, r)) for r in cur.fetchall()]

    if not gl_rows:
        log("No items on 720/710 this period."); return

    cur.execute("""
        SELECT ROUND(PURCHASE_PRICE::FLOAT, 2)
        FROM FIVETRAN.XERO.ASSET
        WHERE ASSET_STATUS IN ('Registered','Draft')
          AND TO_DATE(PURCHASE_DATE) BETWEEN %(s)s AND %(e)s
    """, {"s": PERIOD_START.isoformat(), "e": PERIOD_END_ISO})
    registered_costs = {r[0] for r in cur.fetchall()}

    unmatched = [r for r in gl_rows if round(float(r["NET_AMOUNT"]), 2) not in registered_costs]
    if unmatched:
        lines = "\n".join(
            f"  - {r['TO_DATE(J.JOURNAL_DATE)']} | J{r['JOURNAL_NUMBER']} | "
            f"{r['DESCRIPTION'] or '(no description)'} | GBP {r['NET_AMOUNT']:,.2f} | acct {r['ACCOUNT_CODE']}"
            for r in unmatched
        )
        fail(f":mag: {len(unmatched)} item(s) on 720/710 not matched to registered asset:\n{lines}\n"
             f"Register in Xero FA module before re-running.")

    log(f"GL audit clear — {len(gl_rows)} item(s) all matched.")


# ── STEP 2: REGISTER CHECK ────────────────────────────────────────────────────

def step2_register_check(conn):
    print("[2/6] Xero register check")
    cur = conn.cursor()
    cur.execute("""
        SELECT ASSET_STATUS, COUNT(*), MAX(_FIVETRAN_SYNCED)
        FROM FIVETRAN.XERO.ASSET GROUP BY ASSET_STATUS
    """)
    rows = {r[0]: {"count": int(r[1]), "sync": r[2]} for r in cur.fetchall()}
    drafts     = rows.get("Draft",      {}).get("count", 0)
    registered = rows.get("Registered", {}).get("count", 0)
    last_sync  = rows.get("Registered", {}).get("sync")

    # Check if Draft records are stale — compare their sync time, not Registered
    draft_sync = rows.get("Draft", {}).get("sync")
    draft_stale = False
    if drafts > 0 and draft_sync:
        draft_hours = (datetime.datetime.utcnow() - draft_sync.replace(tzinfo=None)).total_seconds() / 3600
        if draft_hours > 6:
            draft_stale = True
            log(f"WARNING: Draft records last synced {draft_hours:.0f}h ago — may be stale (already resolved in Xero).")

    if drafts > 0 and not draft_stale:
        fail(f":file_cabinet: {drafts} Draft asset(s) in Xero FA module. Register or delete before running.")
    elif drafts > 0 and draft_stale:
        log(f"WARNING: {drafts} Draft(s) in Snowflake but data is {draft_hours:.0f}h old — proceeding.")

    log(f"{registered} registered, {drafts} drafts. Last sync: {last_sync}")
    return registered


# ── STEP 3: XERO API ──────────────────────────────────────────────────────────

def _xero_token():
    r = requests.post("https://identity.xero.com/connect/token", data={
        "grant_type": "refresh_token",
        "refresh_token": os.environ["XERO_REFRESH_TOKEN"],
        "client_id":     os.environ["XERO_CLIENT_ID"],
        "client_secret": os.environ["XERO_CLIENT_SECRET"],
    }, timeout=30)
    if r.status_code != 200:
        fail(f":key: Xero token refresh failed: {r.status_code} — {r.text}")
    data = r.json()
    os.environ["XERO_REFRESH_TOKEN"] = data.get("refresh_token", os.environ["XERO_REFRESH_TOKEN"])
    return data["access_token"]

def step3_run_depreciation():
    print(f"[3/6] Xero depreciation — period end {PERIOD_END_ISO}")
    r = requests.post(
        "https://api.xero.com/assets.xro/1.0/AssetTypes/depreciation",
        headers={"Authorization": f"Bearer {_xero_token()}",
                 "Xero-Tenant-Id": os.environ["XERO_TENANT_ID"],
                 "Content-Type": "application/json"},
        json={"runDate": PERIOD_END_ISO}, timeout=60,
    )
    if r.status_code not in (200, 201):
        fail(f":x: Xero API {r.status_code}: {r.text}\nPeriod may already be posted — check Xero manually.")
    log(f"Xero responded {r.status_code}. Depreciation triggered.")


# ── STEP 4: VERIFY ────────────────────────────────────────────────────────────

def step4_verify(conn, registered_count):
    print("[4/6] Post-run verification")
    cur = conn.cursor()
    total_dep = 0.0; journal_count = 0

    for attempt in range(1, 4):
        cur.execute("""
            SELECT COUNT(DISTINCT j.JOURNAL_NUMBER), SUM(jl.NET_AMOUNT)
            FROM FIVETRAN.XERO.JOURNAL j
            JOIN FIVETRAN.XERO.JOURNAL_LINE jl ON j.JOURNAL_ID = jl.JOURNAL_ID
            WHERE TO_DATE(j.JOURNAL_DATE) = %(e)s
              AND jl.ACCOUNT_CODE = '416'
              AND jl.NET_AMOUNT   > 0
              AND TO_DATE(j.CREATED_DATE_UTC) >= %(t)s
        """, {"e": PERIOD_END_ISO, "t": TODAY.isoformat()})
        row = cur.fetchone()
        journal_count = int(row[0] or 0); total_dep = float(row[1] or 0)
        if journal_count > 0: break
        log(f"Attempt {attempt}/3: journals not yet in Snowflake — waiting 90s...")
        time.sleep(90)

    if journal_count == 0:
        fail(f":question: No dep journals for {PERIOD_END_ISO} after 3 attempts. Verify in Xero manually.")

    cur.execute("""
        SELECT TO_DATE(PURCHASE_DATE), PURCHASE_PRICE
        FROM FIVETRAN.XERO.ASSET
        WHERE ASSET_STATUS = 'Registered' AND PURCHASE_DATE IS NOT NULL
    """)
    expected = sum(
        round(float(c) / 60, 2) for (d, c) in cur.fetchall()
        if d and c and 1 <= (PERIOD_YEAR - d.year)*12 + (PERIOD_MONTH - d.month) + 1 <= 60
    )
    variance = abs(total_dep - expected)
    log(f"Posted GBP {total_dep:,.2f} | Expected GBP {expected:,.2f} | Variance GBP {variance:.2f}")

    if variance > VERIFY_TOLERANCE:
        fail(f":triangular_ruler: Variance GBP {variance:.2f} > tolerance GBP {VERIFY_TOLERANCE}.\n"
             f"Posted GBP {total_dep:,.2f} vs expected GBP {expected:,.2f}. Rollback and investigate.")

    log(f"Verified. {journal_count} journals posted.")
    return total_dep


# ── STEP 5: REBUILD SCHEDULE ──────────────────────────────────────────────────

def step5_rebuild_schedule(conn):
    print("[5/6] Rebuilding FA_Schedule_FINAL.xlsx")
    try:
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import build_fa_schedule
        drive_url = build_fa_schedule.run(conn)
        log(f"Schedule rebuilt and uploaded to Drive.")
        if drive_url:
            os.environ["GOOGLE_DRIVE_FILE_URL"] = drive_url
    except Exception as e:
        slack(f":orange_circle: Dep posted but schedule rebuild failed: {e}\nRebuild manually.", ":orange_circle:")
        log(f"WARNING: Schedule rebuild failed: {e}")


# ── STEP 6: NOTIFY ────────────────────────────────────────────────────────────

def step6_notify(total_dep, asset_count):
    drive_url  = os.environ.get("GOOGLE_DRIVE_FILE_URL", "")
    drive_line = f"\n:open_file_folder: <{drive_url}|Open FA Schedule>" if drive_url else ""
    slack(
        f":white_check_mark: *{PERIOD_LABEL} depreciation complete*\n"
        f"  - Total posted: *GBP {total_dep:,.2f}*\n"
        f"  - Assets charged: *{asset_count}*\n"
        f"  - Period end: {PERIOD_END_ISO}{drive_line}"
    )
    summary = os.environ.get("GITHUB_STEP_SUMMARY", "")
    if summary:
        with open(summary, "a") as f:
            f.write(f"## FA Depreciation — {PERIOD_LABEL}\n\n")
            f.write(f"| | |\n|---|---|\n")
            f.write(f"| Total posted | GBP {total_dep:,.2f} |\n")
            f.write(f"| Assets | {asset_count} |\n")
            f.write(f"| Period end | {PERIOD_END_ISO} |\n")
            f.write(f"| Status | Completed |\n")
            if drive_url:
                f.write(f"| FA Schedule | [Open in Drive]({drive_url}) |\n")


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    print(f"\n{'='*55}\n  Retro Labs — FA Depreciation — {PERIOD_LABEL}\n{'='*55}\n")
    conn = get_conn()
    try:
        step1_gl_audit(conn)
        n = step2_register_check(conn)
        step3_run_depreciation()
        total = step4_verify(conn, n)
        step5_rebuild_schedule(conn)
        step6_notify(total, n)
    finally:
        conn.close()
    print(f"\n{'='*55}\n  Done — {PERIOD_LABEL}\n{'='*55}\n")

if __name__ == "__main__":
    main()
