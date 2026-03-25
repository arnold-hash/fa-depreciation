# FA Depreciation Automation — Runbook
**Retro Labs Ltd** | Head of Finance: Arnold Mavunga | Last updated: March 2026

---

## What this does

Runs automatically on the **1st of every month at 08:00 UTC** via GitHub Actions. No manual intervention required once set up — unless the audit catches an issue.

The workflow:
1. Audits the Snowflake GL (accounts 720/710) for unregistered assets
2. Checks the Xero FA register is clean (0 drafts, Fivetran sync current)
3. Triggers depreciation via the Xero API for the prior month
4. Verifies journals posted in Snowflake within ±£10 tolerance
5. **Rebuilds `FA_Schedule_FINAL.xlsx` from live Snowflake data and uploads to Google Drive**
6. Posts a Slack summary with the period total and Drive link

---

## Repo structure

```
.github/
  workflows/
    fa_dep.yml              ← Cron schedule + job definition
scripts/
  fa_depreciation.py        ← Main orchestration (steps 1–6)
  build_fa_schedule.py      ← Excel rebuild + Google Drive upload
FA_DEPRECIATION_RUNBOOK.md  ← This file
```

---

## One-time setup

### 1. GitHub Actions secrets

Go to **Settings → Secrets and variables → Actions** in your repo and add:

| Secret | Where to find it |
|---|---|
| `SNOWFLAKE_ACCOUNT` | Snowflake admin → account identifier (e.g. `abc12345.eu-west-1`) |
| `SNOWFLAKE_USER` | Your Snowflake username |
| `SNOWFLAKE_PASSWORD` | Your Snowflake password |
| `SNOWFLAKE_DATABASE` | `FIVETRAN` |
| `SNOWFLAKE_WAREHOUSE` | Your warehouse name |
| `SNOWFLAKE_SCHEMA` | `XERO` |
| `XERO_CLIENT_ID` | Xero developer portal → your app |
| `XERO_CLIENT_SECRET` | Xero developer portal → your app |
| `XERO_TENANT_ID` | See step 2 below |
| `XERO_REFRESH_TOKEN` | See step 2 below |
| `SLACK_WEBHOOK_URL` | Slack → Apps → Incoming Webhooks |
| `GOOGLE_DRIVE_FOLDER_ID` | Google Drive → folder URL → ID after `/folders/` |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | See step 3 below |

---

### 2. Xero OAuth2 app setup

This is required for step 3 (triggering depreciation via API). One-time only.

1. Go to [developer.xero.com](https://developer.xero.com) → **New app**
2. App type: **Web app**
3. Redirect URI: `https://localhost` (for initial token exchange only)
4. Copy `Client ID` and `Client Secret` → add to GitHub secrets
5. Exchange for tokens using the Xero OAuth2 flow:
   - Scope required: `assets`
   - Authorise the app against your Retro Labs Xero org
   - You will receive an `access_token` and `refresh_token`
6. Add `XERO_REFRESH_TOKEN` to GitHub secrets
7. To find `XERO_TENANT_ID`: call `GET https://api.xero.com/connections` with your access token — returns the tenant ID for Retro Labs Ltd

> **Note:** The refresh token is long-lived but not infinite. The script automatically rotates it on each run. If it ever expires (after ~60 days of no runs), you will need to re-authorise manually and update the `XERO_REFRESH_TOKEN` secret.

---

### 3. Google Drive service account

Required for the schedule rebuild upload (step 5).

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a new project (or use an existing one)
3. Enable the **Google Drive API**
4. **IAM → Service Accounts → Create** → download the JSON key file
5. Add `GOOGLE_SERVICE_ACCOUNT_JSON` secret = the full contents of the JSON file
6. In Google Drive, **share your FA schedule folder** with the service account email (found in the JSON as `client_email`) — give **Editor** access
7. Add `GOOGLE_DRIVE_FOLDER_ID` = the ID from the folder URL

---

## How the schedule rebuild works (step 5)

`build_fa_schedule.py` is the key step for keeping your overview file current. Every month it:

- Pulls **all assets** from `FIVETRAN.XERO.ASSET` live
- Pulls **all posted dep journals** from `FIVETRAN.XERO.JOURNAL` (account 721/711) live
- Rebuilds all 4 tabs from scratch:
  - **Summary** — KPIs, last 6 periods table, auto-generated date stamp
  - **Asset Register** — all 91+ assets with live book values from Xero
  - **Monthly Dep Schedule** — full history Oct 2021 → Jul 2026 (expands automatically as new assets are added)
  - **Xero Reconciliation** — Xero actuals vs schedule, last 12 periods
- Uploads to the same Google Drive file (overwrite, not duplicate)
- Posts the Drive link in the Slack notification

The file is also saved as a **GitHub Actions workflow artifact** (90-day retention) as a backup, accessible from the Actions run page.

---

## Manual trigger

You can run the workflow manually at any time from GitHub:

1. Go to **Actions → FA Depreciation — Month-End**
2. Click **Run workflow**
3. Leave `dry_run` as `false` to run the full flow

Useful when:
- You need to re-run after fixing a GL issue mid-month
- Testing after initial setup
- Re-running the schedule rebuild without a full depreciation run (edit `fa_depreciation.py` to call only `step5_rebuild_schedule`)

---

## What happens when it fails

The workflow sends a **Slack alert to `#finance-ops`** (or wherever your webhook posts) with the specific failure reason, then exits. No partial posting — if any step fails before step 3, Xero is never touched.

| Failure | Likely cause | Action |
|---|---|---|
| Step 1 — GL audit | New asset on 720/710 not registered in FA module | Register in Xero FA module, then re-trigger |
| Step 1 — GL audit | Asset registered at wrong cost | Correct cost in Xero FA module, re-trigger |
| Step 2 — Draft assets | Someone saved a draft FA this month | Register or delete the draft in Xero |
| Step 2 — Fivetran stale | Pipeline hasn't run in >25h | Check Fivetran dashboard, force sync, re-trigger |
| Step 3 — Xero API | Dep already posted | Check Xero — may be a duplicate trigger, no action needed |
| Step 3 — Xero API 401 | Refresh token expired | Re-authorise Xero OAuth2 app, update secret |
| Step 4 — Verification | Variance >£10 | Check Xero journals manually, rollback if needed |
| Step 5 — Schedule rebuild | Drive credentials issue | Check service account permissions, rebuild manually |

---

## Outstanding manual tasks (as at March 2026)

| Priority | Item |
|---|---|
| 🔴 | FA-0105 Neil Shah cost: registered at £1,249.17, actual £1,499.00. Correct in Xero FA module. |
| 🟡 | FA-0028 disposal journal: MacBook Air M3 Space Grey. Verify NBV write-off posted correctly. |
| 🟡 | Assets FA-0001–FA-0013 (2021–22): confirm current assignees with IT/Anna. |
| 🟡 | Xero OAuth2 app: set up before first automated run (required for step 3). |

---

## Key technical details

**Depreciation method:** Straight-line, 5 years (60 months), full month averaging. Monthly charge = cost ÷ 60.

**Capitalisation threshold:** All IT hardware and equipment on accounts 720/710 is capitalised regardless of value.

**Xero accounts:**
- `720` Computer Equipment (BS) → accumulated dep `721`
- `710` Office Equipment (BS) → accumulated dep `711`
- `416` Depreciation Expense (P&L)

**Snowflake tables used:**
- `FIVETRAN.XERO.ASSET` — FA register
- `FIVETRAN.XERO.ASSET_TYPE` — asset type lookup
- `FIVETRAN.XERO.JOURNAL` + `FIVETRAN.XERO.JOURNAL_LINE` — posted journals

**Verification tolerance:** ±£10 (accounts for Xero's accumulated book value rounding vs fresh cost÷60 calculation).

---

## Dependencies (auto-installed by workflow)

```
snowflake-connector-python
python-dateutil
openpyxl
requests
google-auth
google-api-python-client
python-dotenv
```

For local development, create a `.env` file with the secrets and run:
```bash
pip install -r requirements.txt
cd scripts
python fa_depreciation.py
```

---

*Runbook maintained by Finance. For questions contact Arnold Mavunga or raise a Linear FIN ticket.*
