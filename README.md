# Intercom → Assembled Queue Sync

Automatically syncs agent queue assignments from Intercom to Assembled every 6 hours via GitHub Actions.

## How it works

1. Fetches all **teams** from Intercom (e.g. "Billing", "Tech Support")
2. Fetches all **admins** from Intercom, each with their team memberships
3. Fetches all **queues** from Assembled
4. Fetches all **people** from Assembled
5. Matches teams → queues by **name** (case-insensitive)
6. Matches agents between systems by **email address**
7. Updates each person in Assembled with their correct queue assignments

No hardcoded mappings needed — as long as your Intercom team names and Assembled queue names match, it just works.

---

## Setup (one-time, ~5 minutes)

### 1. Get your API keys

**Intercom:**
- Go to your [Intercom Developer Hub](https://developers.intercom.com/)
- Create an app or use an existing one
- Copy the **Access Token** from the Authentication section

**Assembled:**
- Go to **Settings → API** in Assembled
- Generate a new API key (starts with `sk_live_`)

### 2. Add secrets to GitHub

In your GitHub repository:

1. Go to **Settings → Secrets and variables → Actions**
2. Click **New repository secret** and add:

| Secret name         | Value                        |
|---------------------|------------------------------|
| `INTERCOM_TOKEN`    | Your Intercom access token   |
| `ASSEMBLED_API_KEY` | Your Assembled API key       |

> ⚠️ Never paste API keys into the code or commit them to the repo. GitHub Secrets are encrypted and only exposed to Actions runs.

### 3. Add the files to your repo

```
your-repo/
├── sync.py
└── .github/
    └── workflows/
        └── sync.yml
```

Push to your `main` branch and the schedule will activate automatically.

### 4. Test it manually

- Go to **Actions** tab in GitHub
- Select **Intercom → Assembled Queue Sync**
- Click **Run workflow**
- Watch the logs to confirm everything is matching correctly

---

## Checking the logs

Each run prints a clear summary like:

```
2024-01-15 06:00:01 [INFO] Fetching teams from Intercom...
2024-01-15 06:00:01 [INFO]   Found 3 Intercom team(s): ['Billing', 'Tech Support', 'General']
2024-01-15 06:00:02 [INFO]   ✓ Matched: Intercom team 'Billing' → Assembled queue 'Billing'
2024-01-15 06:00:02 [INFO]   ✓ Matched: Intercom team 'Tech Support' → Assembled queue 'Tech Support'
2024-01-15 06:00:02 [WARNING]   ⚠ 1 Intercom team(s) had no matching Assembled queue: ['General']
2024-01-15 06:00:03 [INFO]   ↻ Updating Jane Smith (jane@example.com): [] → ['uuid-billing']
2024-01-15 06:00:03 [INFO]   – Joe Bloggs (joe@example.com): no change needed
...
── Sync complete ────────────────────────────────
  Updated : 2
  Skipped : 5 (already correct)
  No match: 0 (Intercom admin not found in Assembled)
─────────────────────────────────────────────────
```

---

## Troubleshooting

**"No matching Assembled queue" warning**
The Intercom team name and Assembled queue name don't match exactly. Check for typos or differences in capitalisation — the match is case-insensitive but otherwise exact.

**"No Assembled person found" warning**
An Intercom admin's email doesn't exist in Assembled. Either add them to Assembled or they can be safely ignored (e.g. a bot account).

**API errors (401)**
Your API key is invalid or expired. Re-generate and update the GitHub Secret.

**API errors (429)**
Rate limit hit. This is unlikely with a 6-hour schedule but if running manually many times, wait a minute and retry.

---

## Adjusting the schedule

Edit the `cron` line in `.github/workflows/sync.yml`:

```yaml
- cron: "0 */6 * * *"   # every 6 hours (default)
- cron: "0 */4 * * *"   # every 4 hours
- cron: "0 8 * * 1-5"   # 8am UTC on weekdays only
```

Use [crontab.guru](https://crontab.guru) to build custom schedules.
