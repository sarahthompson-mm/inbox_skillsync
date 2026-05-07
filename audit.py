#!/usr/bin/env python3
"""
Intercom ↔ Assembled Agent Audit
==================================
Compares agents between Intercom and Assembled using the Intercom platform ID
stored in Assembled's platforms.intercom field for accurate matching.

Required environment variables:
  INTERCOM_TOKEN     - Intercom API Bearer token
  ASSEMBLED_API_KEY  - Assembled API key (sk_live_...)

Output:
  agent_audit.xlsx
"""

import os
import sys
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

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

GREEN  = "C6EFCE"
RED    = "FFC7CE"
YELLOW = "FFEB9C"
GREY   = "D9D9D9"
BLUE   = "BDD7EE"


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


def fetch_intercom_admins():
    print("Fetching Intercom admins...")
    data = intercom_get("/admins")
    admins = {}
    for admin in data.get("admins", []):
        admins[str(admin["id"])] = {
            "name":     admin.get("name", ""),
            "email":    admin.get("email", "").lower().strip(),
            "team_ids": [str(t) for t in admin.get("team_ids", [])],
        }
    print(f"  Found {len(admins)} Intercom admins")
    return admins


def fetch_intercom_teams():
    print("Fetching Intercom teams...")
    data = intercom_get("/teams")
    teams = {}
    for team in data.get("teams", []):
        teams[str(team["id"])] = team["name"]
    print(f"  Found {len(teams)} Intercom teams")
    return teams


def fetch_assembled_people():
    print("Fetching Assembled people...")
    all_people = assembled_get_all("/people", "people")
    people = {}
    no_id = []
    for person in all_people.values():
        if person.get("deleted"):
            continue
        intercom_id = person.get("platforms", {}).get("intercom", "")
        name = f"{person.get('first_name','')} {person.get('last_name','')}".strip()
        if intercom_id:
            people[str(intercom_id)] = {
                "id":     person["id"],
                "name":   name,
                "email":  person.get("email", ""),
                "queues": person.get("queues", []),
            }
        else:
            no_id.append({"name": name, "email": person.get("email", "")})
    print(f"  Found {len(people)} Assembled people with Intercom platform ID")
    print(f"  {len(no_id)} Assembled people have no Intercom platform ID")
    return people, no_id


def fetch_assembled_queues():
    print("Fetching Assembled queues...")
    all_queues = assembled_get_all("/queues", "queues")
    queues = {}
    for queue in all_queues.values():
        queues[queue["id"]] = queue["name"]
    print(f"  Found {len(queues)} Assembled queues")
    return queues


def header_row(sheet, headers, colour=GREY):
    sheet.append(headers)
    for cell in sheet[sheet.max_row]:
        cell.font      = Font(bold=True, name="Arial")
        cell.fill      = PatternFill("solid", start_color=colour)
        cell.alignment = Alignment(horizontal="center")


def colour_row(sheet, row_num, colour):
    for cell in sheet[row_num]:
        cell.fill = PatternFill("solid", start_color=colour)


def set_column_widths(sheet, widths):
    for i, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(i)].width = width


def build_report(intercom_admins, intercom_teams, assembled_people, assembled_queues, no_intercom_id):
    wb = Workbook()

    bot_keywords = ["intercom.io", "facebookbot", "gmail.com", "operator+"]

    matched        = []
    ic_only        = []
    ic_bots        = []
    assembled_only = []

    for admin_id, admin in intercom_admins.items():
        is_bot = any(k in admin["email"] for k in bot_keywords)
        if is_bot:
            ic_bots.append(admin)
        elif admin_id in assembled_people:
            matched.append({**admin, "assembled": assembled_people[admin_id]})
        else:
            ic_only.append(admin)

    for intercom_id, person in assembled_people.items():
        if intercom_id not in intercom_admins:
            assembled_only.append(person)

    # ── Sheet 1: Summary ──────────────────────────────────────────────────────
    ws_summary = wb.active
    ws_summary.title = "Summary"
    ws_summary.append(["Intercom ↔ Assembled Agent Audit"])
    ws_summary["A1"].font = Font(bold=True, size=14, name="Arial")
    ws_summary.append([])

    summary_data = [
        ["Metric",                                          "Count"],
        ["✅ Matched via Intercom platform ID",             len(matched)],
        ["⚠️ In Intercom only (no match in Assembled)",     len(ic_only)],
        ["🤖 Bots / system accounts (Intercom)",            len(ic_bots)],
        ["⚠️ In Assembled only (no match in Intercom)",     len(assembled_only)],
        ["❌ Assembled people missing Intercom ID",         len(no_intercom_id)],
        ["📊 Total Intercom admins",                        len(intercom_admins)],
        ["📊 Total Assembled people (active)",              len(assembled_people)],
    ]

    for i, row in enumerate(summary_data):
        ws_summary.append(row)
        if i == 0:
            for cell in ws_summary[ws_summary.max_row]:
                cell.font = Font(bold=True, name="Arial")
                cell.fill = PatternFill("solid", start_color=GREY)
        elif i == 1:
            colour_row(ws_summary, ws_summary.max_row, GREEN)
        elif i in (2, 4, 5):
            colour_row(ws_summary, ws_summary.max_row, YELLOW)

    set_column_widths(ws_summary, [50, 10])

    # ── Sheet 2: Matched ──────────────────────────────────────────────────────
    ws_matched = wb.create_sheet("✅ Matched")
    header_row(ws_matched, ["Name", "Email", "Intercom ID", "Intercom Teams", "Assembled Queues"], GREEN)

    for m in sorted(matched, key=lambda x: x["name"]):
        team_names  = ", ".join(intercom_teams.get(tid, tid) for tid in m["team_ids"]) or "No team"
        queue_names = ", ".join(assembled_queues.get(qid, qid) for qid in m["assembled"]["queues"]) or "No queue"
        ws_matched.append([m["name"], m["email"], m["assembled"]["id"], team_names, queue_names])

    set_column_widths(ws_matched, [25, 35, 15, 40, 40])

    # ── Sheet 3: Intercom Only ────────────────────────────────────────────────
    ws_ic = wb.create_sheet("⚠️ Intercom Only")
    header_row(ws_ic, ["Name", "Intercom Email", "Intercom Teams", "Action"], YELLOW)

    for admin in sorted(ic_only, key=lambda x: x["name"]):
        team_names = ", ".join(intercom_teams.get(tid, tid) for tid in admin["team_ids"]) or "No team"
        ws_ic.append([admin["name"], admin["email"], team_names, "Add to Assembled & set Intercom platform ID"])

    set_column_widths(ws_ic, [25, 35, 40, 40])

    # ── Sheet 4: Assembled Only ───────────────────────────────────────────────
    ws_as = wb.create_sheet("⚠️ Assembled Only")
    header_row(ws_as, ["Name", "Assembled Email", "Assembled Queues", "Action"], YELLOW)

    for person in sorted(assembled_only, key=lambda x: x["name"]):
        queue_names = ", ".join(assembled_queues.get(qid, qid) for qid in person["queues"]) or "No queue"
        ws_as.append([person["name"], person["email"], queue_names, "Add to Intercom OR check platform ID"])

    set_column_widths(ws_as, [25, 35, 40, 38])

    # ── Sheet 5: Missing Intercom ID ──────────────────────────────────────────
    ws_noid = wb.create_sheet("❌ Missing Intercom ID")
    header_row(ws_noid, ["Name", "Assembled Email", "Action"], RED)

    for person in sorted(no_intercom_id, key=lambda x: x["name"]):
        ws_noid.append([person["name"], person["email"], "Set Intercom platform ID in Assembled"])

    set_column_widths(ws_noid, [25, 35, 38])

    for ws in [ws_matched, ws_ic, ws_as, ws_noid]:
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if not cell.font or not cell.font.bold:
                    cell.font = Font(name="Arial")

    return wb, len(matched), len(ic_only), len(assembled_only), len(no_intercom_id)


def main():
    missing = [v for v in ("INTERCOM_TOKEN", "ASSEMBLED_API_KEY") if not os.environ.get(v)]
    if missing:
        print(f"Missing environment variables: {', '.join(missing)}")
        sys.exit(1)

    print("═══════════════════════════════════════")
    print("  Intercom ↔ Assembled Agent Audit")
    print("═══════════════════════════════════════")

    intercom_admins  = fetch_intercom_admins()
    intercom_teams   = fetch_intercom_teams()
    assembled_people, no_intercom_id = fetch_assembled_people()
    assembled_queues = fetch_assembled_queues()

    print("\nBuilding report...")
    wb, matched, ic_only, as_only, missing_id = build_report(
        intercom_admins, intercom_teams, assembled_people, assembled_queues, no_intercom_id
    )

    output = "agent_audit.xlsx"
    wb.save(output)

    print(f"\n── Audit complete ──────────────────────")
    print(f"  ✅ Matched:                {matched}")
    print(f"  ⚠️  Intercom only:          {ic_only}")
    print(f"  ⚠️  Assembled only:         {as_only}")
    print(f"  ❌ Missing Intercom ID:    {missing_id}")
    print(f"  📁 Report saved to:        {output}")
    print(f"────────────────────────────────────────")


if __name__ == "__main__":
    main()
