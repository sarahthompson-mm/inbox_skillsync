#!/usr/bin/env python3
"""
Intercom → Assembled Queue Sync
================================
Fetches team assignments from Intercom and syncs them to Assembled queues.
Agents are matched between systems using the Intercom platform ID stored
in Assembled's platforms.intercom field.
Teams/Queues are matched by name (case-insensitive).

Required environment variables:
  INTERCOM_TOKEN     - Intercom API Bearer token
  ASSEMBLED_API_KEY  - Assembled API key (sk_live_...)
"""

import os
import sys
import logging
import requests

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

INTERCOM_TOKEN    = os.environ.get("INTERCOM_TOKEN")
ASSEMBLED_API_KEY = os.environ.get("ASSEMBLED_API_KEY")

INTERCOM_BASE  = "https://api.intercom.io"
ASSEMBLED_BASE = "https://api.assembledhq.com/v0"

INTERCOM_HEADERS = {
    "Authorization": f"Bearer {INTERCOM_TOKEN}",
    "Intercom-Version": "2.11",
    "Accept": "application/json",
}
ASSEMBLED_AUTH = (ASSEMBLED_API_KEY, "")


def intercom_get(path):
    r = requests.get(f"{INTERCOM_BASE}{path}", headers=INTERCOM_HEADERS, timeout=30)
    r.raise_for_status()
    return r.json()


def assembled_get(path, params=None):
    r = requests.get(f"{ASSEMBLED_BASE}{path}", auth=ASSEMBLED_AUTH, params=params, timeout=30)
    r.raise_for_status()
    return r.json()


def assembled_get_all(path, key):
    """Fetch all pages from a paginated Assembled endpoint."""
    results = {}
    offset = 0
    limit = 100
    while True:
        data = assembled_get(path, params={"limit": limit, "offset": offset})
        page = data.get(key, {})
        if isinstance(page, dict):
            results.update(page)
        else:
            for item in page:
                results[item["id"]] = item
        total = data.get("total", 0)
        offset += limit
        if offset >= total:
            break
    return results


def assembled_patch(path, payload):
    r = requests.patch(f"{ASSEMBLED_BASE}{path}", auth=ASSEMBLED_AUTH, json=payload, timeout=30)
    r.raise_for_status()
    return r.json()


def fetch_intercom_teams():
    log.info("Fetching teams from Intercom...")
    data = intercom_get("/teams")
    teams = {}
    for team in data.get("teams", []):
        teams[str(team["id"])] = {
            "name":      team["name"],
            "admin_ids": [str(a) for a in team.get("admin_ids", [])],
        }
    log.info(f"  Found {len(teams)} Intercom team(s)")
    return teams


def fetch_intercom_admins():
    log.info("Fetching admins from Intercom...")
    data = intercom_get("/admins")
    admins = {}
    for admin in data.get("admins", []):
        admins[str(admin["id"])] = {
            "name":     admin.get("name", ""),
            "email":    admin.get("email", "").lower().strip(),
            "team_ids": [str(t) for t in admin.get("team_ids", [])],
        }
    log.info(f"  Found {len(admins)} Intercom admin(s)")
    for aid, a in list(admins.items())[:3]:
        log.info(f"  Sample Intercom admin ID: '{aid}' -> {a['name']}")
    return admins


def fetch_assembled_queues():
    log.info("Fetching queues from Assembled...")
    all_queues = assembled_get_all("/queues", "queues")
    queues = {}
    for queue in all_queues.values():
        queues[queue["name"].lower().strip()] = queue["id"]
    log.info(f"  Found {len(queues)} Assembled queue(s)")
    return queues


def fetch_assembled_people():
    log.info("Fetching people from Assembled...")
    all_people = assembled_get_all("/people", "people")
    # Keyed by Intercom platform ID for direct matching
    people = {}
    no_intercom_id = 0
    for person in all_people.values():
        if person.get("deleted"):
            continue
        intercom_id = person.get("platforms", {}).get("intercom", "")
        if intercom_id:
            people[str(intercom_id)] = {
                "id":     person["id"],
                "name":   f"{person.get('first_name','')} {person.get('last_name','')}".strip(),
                "email":  person.get("email", ""),
                "queues": person.get("queues", []),
            }
        else:
            no_intercom_id += 1
    log.info(f"  Found {len(people)} Assembled people with Intercom platform ID")
    for iid, p in list(people.items())[:3]:
        log.info(f"  Sample Assembled Intercom ID: '{iid}' -> {p['name']}")
    if no_intercom_id:
        log.warning(f"  ⚠ {no_intercom_id} Assembled people have no Intercom platform ID set")
    return people


def build_target_queues(intercom_teams, intercom_admins, assembled_queues):
    team_to_queue = {}
    unmatched_teams = []

    for team_id, team in intercom_teams.items():
        team_name_lower = team["name"].lower().strip()
        if team_name_lower in assembled_queues:
            team_to_queue[team_id] = assembled_queues[team_name_lower]
            log.info(f"  ✓ Matched: '{team['name']}'")
        else:
            unmatched_teams.append(team["name"])

    if unmatched_teams:
        log.warning(
            f"  ⚠ {len(unmatched_teams)} Intercom team(s) had no matching Assembled queue: {unmatched_teams}"
        )

    # Build per-admin target queues keyed by Intercom admin ID
    admin_to_queues = {}
    for admin_id, admin in intercom_admins.items():
        target = [
            team_to_queue[tid]
            for tid in admin["team_ids"]
            if tid in team_to_queue
        ]
        admin_to_queues[admin_id] = target

    return admin_to_queues


def sync_to_assembled(admin_to_target_queues, assembled_people):
    updated  = 0
    skipped  = 0
    no_match = 0

    for intercom_id, target_queues in admin_to_target_queues.items():
        person = assembled_people.get(intercom_id)

        if not person:
            no_match += 1
            continue

        current_queues = sorted(person["queues"])
        desired_queues = sorted(list(set(target_queues)))

        if current_queues == desired_queues:
            log.info(f"  – {person['name']}: no change needed")
            skipped += 1
            continue

        log.info(f"  ↻ Updating {person['name']}: {current_queues} → {desired_queues}")
        assembled_patch(f"/people/{person['id']}", {"queues": desired_queues})
        updated += 1

    log.info("")
    log.info("── Sync complete ────────────────────────────────")
    log.info(f"  Updated : {updated}")
    log.info(f"  Skipped : {skipped} (already correct)")
    log.info(f"  No match: {no_match} (Intercom admin not found in Assembled)")
    log.info("─────────────────────────────────────────────────")


def main():
    missing = [v for v in ("INTERCOM_TOKEN", "ASSEMBLED_API_KEY") if not os.environ.get(v)]
    if missing:
        log.error(f"Missing required environment variable(s): {', '.join(missing)}")
        sys.exit(1)

    log.info("═══════════════════════════════════════════════")
    log.info("  Intercom → Assembled Queue Sync")
    log.info("═══════════════════════════════════════════════")

    try:
        intercom_teams   = fetch_intercom_teams()
        intercom_admins  = fetch_intercom_admins()
        assembled_queues = fetch_assembled_queues()
        assembled_people = fetch_assembled_people()

        log.info("")
        log.info("Matching teams to queues...")
        admin_to_target_queues = build_target_queues(
            intercom_teams, intercom_admins, assembled_queues
        )

        log.info("")
        log.info("Syncing to Assembled...")
        sync_to_assembled(admin_to_target_queues, assembled_people)

    except requests.HTTPError as e:
        log.error(f"API error: {e.response.status_code} {e.response.text}")
        sys.exit(1)
    except Exception as e:
        log.error(f"Unexpected error: {e}")
        raise


if __name__ == "__main__":
    main()
