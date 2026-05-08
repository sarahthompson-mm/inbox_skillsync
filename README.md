# inbox_skillsync

Automated people operations integrations for Marshmallow, connecting Intercom, Assembled and HiBob via GitHub Actions.

---

## Scripts

### `sync.py` — Intercom → Assembled Queue Sync
Runs every 6 hours automatically. Fetches team assignments from Intercom and syncs them to Assembled queues.

- Agents are matched between systems using the **Intercom platform ID** stored in Assembled's `platforms.intercom` field
- Teams and queues are matched by **name** (case-insensitive), with manual overrides for names that will never match
- Protected queues are never removed, even if Intercom has no matching team

### `audit.py` — Intercom ↔ Assembled Agent Audit
Run on demand. Produces an Excel report comparing agents between Intercom and Assembled.

Sheets:
- **Summary** — headline numbers
- **✅ Matched** — agents joined via Intercom platform ID
- **⚠️ Intercom Only** — in Intercom but not Assembled
- **⚠️ Assembled Only** — in Assembled but not Intercom
- **❌ Missing Intercom ID** — Assembled agents with no Intercom platform ID set

### `people_audit.py` — Assembled People Audit
Run on demand. Pulls all active people from Assembled and flags data quality issues.

Sheets:
- **📊 Summary** — role breakdown and headline counts
- **👔 Managers** — all manager-role agents with flags
- **⭐ Team Leads** — team leads separated out
- **🔧 Admins** — admin accounts
- **👤 Agents** — main agent population
- **🏢 Teams** — every team with member count (single-person teams highlighted)

### `hibob_assembled_dryrun.py` — HiBob → Assembled Dry Run
Run on demand. Matches HiBob employees (from `hibob_employees.csv`) to Assembled people by exact email and reports what would change.

Sheets:
- **Summary** — match counts at a glance
- **Matched** — employees joined by email, showing HiBob team vs current Assembled teams
- **Team Changes** — agents whose HiBob team doesn't match Assembled
- **Name Mismatches** — where HiBob and Assembled disagree on someone's name
- **No Match** — HiBob employees with no Assembled record

---

## GitHub Actions Workflows

| Workflow | File | Trigger | What it does |
|---|---|---|---|
| Intercom → Assembled Queue Sync | `sync.yml` | Every 6 hours + manual | Runs `sync.py` |
| Agent Audit Report | `audit.yml` | Manual only | Runs `audit.py`, uploads Excel artifact |
| People Audit Report | `people_audit.yml` | Manual only | Runs `people_audit.py`, uploads Excel artifact |
| HiBob → Assembled Dry Run | `hibob_dryrun.yml` | Manual only | Runs `hibob_assembled_dryrun.py`, uploads Excel artifact |

To trigger a manual run: **Actions tab → select workflow → Run workflow**

To download a report: click into the completed run → scroll to **Artifacts** → download

---

## Setup

### GitHub Secrets required

| Secret | Used by | Description |
|---|---|---|
| `INTERCOM_TOKEN` | sync, audit | Intercom API Bearer token |
| `ASSEMBLED_API_KEY` | sync, audit, people audit, hibob dryrun | Assembled API key (`sk_live_...`) |

Add secrets at: **Settings → Secrets and variables → Actions → New repository secret**

### HiBob CSV

`hibob_assembled_dryrun.py` requires a `hibob_employees.csv` file in the root of the repo. Generate it by running this query in Snowflake and downloading as CSV:

```sql
SELECT DISTINCT
    b.EMPLOYEE_ID       as hibob_id,
    b.EMAIL,
    b.FIRST_NAME,
    b.LAST_NAME,
    b.JOB_TITLE,
    b.TEAM,
    b.LEVEL,
    b.MANAGER_EMAIL,
    a.AGENT_ID          as assembled_id,
    a.INTERCOM_AGENT_ID as intercom_id
FROM ANALYTICS.COPS.BOB_EMPLOYEES b
LEFT JOIN ANALYTICS.COPS.ASSEMBLED__AGENT_ACTIVITIES a
    ON LOWER(b.EMAIL) = LOWER(a.AGENT_EMAIL)
WHERE b.TERMINATED_AT IS NULL
  AND b.EMAIL IS NOT NULL
  AND b.EMPLOYEE_ID IS NOT NULL
ORDER BY b.LAST_NAME, b.FIRST_NAME;
```

---

## Configuration

### Manual queue name overrides (`sync.py`)

For Intercom teams whose names will never match an Assembled queue name:

```python
MANUAL_NAME_OVERRIDES = {
    "aircall fnol": "first party claims - fnol - calls",
}
```

### Protected queues (`sync.py`)

Queues that will never be removed by the sync, even if Intercom has no matching team. Agents keep these alongside whatever Intercom assigns them:

```python
PROTECTED_QUEUE_NAMES = [
    "first party claims - fnol - calls",
]
```

### Team name overrides (`hibob_assembled_dryrun.py`)

For HiBob teams that have been renamed in Assembled:

```python
TEAM_NAME_OVERRIDES = {
    "customer happiness": "customer care",
    "retention":          "customer care",
}
```

---

## Troubleshooting

**Sync updated 0 agents**
Most likely the Intercom platform ID isn't set in Assembled for those agents. Run `audit.py` and check the **❌ Missing Intercom ID** sheet.

**"No matching Assembled queue" warning**
An Intercom team name doesn't match any Assembled queue name. Either fix the name in Assembled, or add a `MANUAL_NAME_OVERRIDES` entry in `sync.py`.

**FNOL agents losing their queue assignment**
The `PROTECTED_QUEUE_NAMES` list in `sync.py` should prevent this. Check it contains `"first party claims - fnol - calls"`.

**API errors (401)**
An API key has expired. Regenerate and update the relevant GitHub Secret.

**HiBob dry run shows 0 matches**
Check the `hibob_employees.csv` is committed to the repo and the email format matches between HiBob and Assembled.

---

## Adjusting the sync schedule

Edit the `cron` line in `.github/workflows/sync.yml`:

```yaml
- cron: "0 */6 * * *"   # every 6 hours (default)
- cron: "0 */4 * * *"   # every 4 hours
- cron: "0 8 * * 1-5"   # 8am UTC weekdays only
```

Use [crontab.guru](https://crontab.guru) to build custom schedules.
