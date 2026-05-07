#!/usr/bin/env python3
"""
Intercom → Assembled Queue Sync
================================
Fetches team assignments from Intercom and syncs them to Assembled queues.
Agents are matched between systems by email address.
Teams/Queues are matched by name (case-insensitive).

Required environment variables:
  INTERCOM_TOKEN     - Intercom API Bearer token
  ASSEMBLED_API_KEY  - Assembled API key (sk_live_...)
"""

import os
import sys
import logging
import requests

# ── Logging ──────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

# ── Config ────────────────────────────────────────────────────────────────────

INTERCOM_TOKEN    = os.environ.get("INTERCOM_TOKEN")
ASSEMBLED_API_KEY = os.environ.get("ASSEMBLED_API_KEY")

INTERCOM_BASE  = "https://api.intercom.io"
ASSEMBLED_BASE = "https://api.assembledhq.com/v0"

INTERCOM_HEADERS = {
    "Authorization": f"Bearer {INTERCOM_TOKEN}",
    "Intercom-Version": "2.11",
    "Accept": "application/json",
}

# Assembled uses HTTP Basic Auth: API key as username, no password
ASSEMBLED_AUTH = (ASSEMBLED_API_KEY, "")

# ── Helpers ───────────────────────────────────────────────────────────────────

def intercom_get(path: str, params: dict = None) -> dict:
    url = f"{INTERCOM_BASE}{path}"
    resp = requests.get(url, headers=INTERCOM_HEADERS, params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()


def assembled_get(path: str, params: dict = None) -> dict:
    url = f"{ASSEMBLED_BASE}{path}"
    resp = requests.get(url, auth=ASSEMBLED_AUTH, params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()


def assembled_patch(path: str, payload: dict) -> dict:
    url = f"{ASSEMBLED_BASE}{path}"
    resp = requests.patch(url, auth=ASSEMBLED_AUTH, json=payload, timeout=30)
    resp.raise_for_status()
    return resp.json()

# ── Step 1: Fetch Intercom teams ──────────────────────────────────────────────

def fetch_intercom_teams() -> dict[str, dict]:
    """
    Returns a dict keyed by team ID:
      { "814865": { "name": "Billing", "admin_ids": [123, 456] }, ... }
    """
    log.info("Fetching teams from Intercom...")
    data = intercom_get("/teams")
    teams = {}
    for team in data.get("teams", []):
        teams[str(team["id"])] = {
            "name": team["name"],
            "admin_ids": [str(a) for a in team.get("admin_ids", [])],
        }
    log.info(f"  Found {len(teams)} Intercom team(s): {[t['name'] for t in teams.values()]}")
    return teams

# ── Step 2: Fetch Intercom admins ─────────────────────────────────────────────

def fetch_intercom_admins() -> dict[str, dict]:
    """
    Returns a dict keyed by admin ID:
      { "123": { "email": "jane@example.com", "name": "Jane", "team_ids": ["814865"] }, ... }
    """
    log.info("Fetching admins from Intercom...")
    data = intercom_get("/admins")
    admins = {}
    for admin in data.get("admins", []):
        admins[str(admin["id"])] = {
            "email": admin.get("email", "").lower().strip(),
            "name":  admin.get("name", ""),
            "team_ids": [str(t) for t in admin.get("team_ids", [])],
        }
    log.info(f"  Found {len(admins)} Intercom admin(s)")
    return admins

# ── Step 3: Fetch Assembled queues ────────────────────────────────────────────

def fetch_assembled_queues() -> dict[str, str]:
    log.info("Fetching queues from Assembled...")
    data = assembled_get("/queues")
    queues = {}
    queue_list = data.get("queues", {}).values()
    for queue in queue_list:
        queues[queue["name"].lower().strip()] = queue["id"]
    log.info(f"  Found {len(queues)} Assembled queue(s): {list(queues.keys())}")
    return queues

# ── Step 4: Fetch Assembled people ────────────────────────────────────────────

def fetch_assembled_people() -> dict[str, dict]:
    """
    Returns a dict keyed by lowercase email:
      { "jane@example.com": { "id": "uuid-xyz", "queues": ["uuid-abc"] }, ... }
    """
    log.info("Fetching people from Assembled...")
    data = assembled_get("/people")
    people = {}
    for person in data.get("people", {}).values()
        email = person.get("email", "").lower().strip()
        if email:
            people[email] = {
                "id":     person["id"],
                "name":   f"{person.get('first_name','')} {person.get('last_name','')}".strip(),
                "queues": person.get("queues", []),
            }
    log.info(f"  Found {len(people)} Assembled person(s)")
    return people

# ── Step 5: Build email → target queue UUIDs map ──────────────────────────────

def build_target_queues(
    intercom_teams: dict,
    intercom_admins: dict,
    assembled_queues: dict,
) -> dict[str, list[str]]:
    """
    For each Intercom admin, work out which Assembled queue UUIDs they should
    have, by matching Intercom team names → Assembled queue names.

    Returns: { "jane@example.com": ["uuid-abc", "uuid-def"], ... }
    """
    # Map Intercom team ID → Assembled queue UUID (by matching names)
    team_to_queue: dict[str, str] = {}
    unmatched_teams: list[str] = []

    for team_id, team in intercom_teams.items():
        team_name_lower = team["name"].lower().strip()
        if team_name_lower in assembled_queues:
            team_to_queue[team_id] = assembled_queues[team_name_lower]
            log.info(f"  ✓ Matched: Intercom team '{team['name']}' → Assembled queue '{team['name']}'")
        else:
            unmatched_teams.append(team["name"])

    if unmatched_teams:
        log.warning(
            f"  ⚠ {len(unmatched_teams)} Intercom team(s) had no matching Assembled queue "
            f"(names must match exactly): {unmatched_teams}"
        )

    # Build per-admin target queue list
    email_to_queues: dict[str, list[str]] = {}
    for admin in intercom_admins.values():
        email = admin["email"]
        if not email:
            continue
        target = [
            team_to_queue[tid]
            for tid in admin["team_ids"]
            if tid in team_to_queue
        ]
        email_to_queues[email] = target

    return email_to_queues

# ── Step 6: Sync to Assembled ─────────────────────────────────────────────────

def sync_to_assembled(
    email_to_target_queues: dict[str, list[str]],
    assembled_people: dict[str, dict],
) -> None:
    updated  = 0
    skipped  = 0
    no_match = 0

    for email, target_queues in email_to_target_queues.items():
        person = assembled_people.get(email)

        if not person:
            log.warning(f"  ⚠ No Assembled person found for Intercom admin: {email}")
            no_match += 1
            continue

        current_queues = sorted(person["queues"])
        desired_queues = sorted(list(set(target_queues)))  # deduplicate

        if current_queues == desired_queues:
            log.info(f"  – {person['name']} ({email}): no change needed")
            skipped += 1
            continue

        log.info(
            f"  ↻ Updating {person['name']} ({email}): "
            f"{current_queues} → {desired_queues}"
        )
        assembled_patch(f"/people/{person['id']}", {"queues": desired_queues})
        updated += 1

    log.info("")
    log.info("── Sync complete ────────────────────────────────")
    log.info(f"  Updated : {updated}")
    log.info(f"  Skipped : {skipped} (already correct)")
    log.info(f"  No match: {no_match} (Intercom admin not found in Assembled)")
    log.info("─────────────────────────────────────────────────")

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    # Validate env vars up front
    missing = [v for v in ("INTERCOM_TOKEN", "ASSEMBLED_API_KEY") if not os.environ.get(v)]
    if missing:
        log.error(f"Missing required environment variable(s): {', '.join(missing)}")
        sys.exit(1)

    log.info("═══════════════════════════════════════════════")
    log.info("  Intercom → Assembled Queue Sync")
    log.info("═══════════════════════════════════════════════")

    try:
        intercom_teams  = fetch_intercom_teams()
        intercom_admins = fetch_intercom_admins()
        assembled_queues = fetch_assembled_queues()
        assembled_people = fetch_assembled_people()

        log.info("")
        log.info("Matching teams to queues and building sync plan...")
        email_to_target_queues = build_target_queues(
            intercom_teams, intercom_admins, assembled_queues
        )

        log.info("")
        log.info("Syncing to Assembled...")
        sync_to_assembled(email_to_target_queues, assembled_people)

    except requests.HTTPError as e:
        log.error(f"API error: {e.response.status_code} {e.response.text}")
        sys.exit(1)
    except Exception as e:
        log.error(f"Unexpected error: {e}")
        raise


if __name__ == "__main__":
    main()
