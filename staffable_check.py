import os
import sys
import requests

ASSEMBLED_API_KEY = os.environ.get("ASSEMBLED_API_KEY")
ASSEMBLED_BASE = "https://api.assembledhq.com/v0"
ASSEMBLED_AUTH = (ASSEMBLED_API_KEY, "")

# Roles that should NOT be staffable
NON_STAFFABLE_ROLE_NAMES = [
    "manager",
    "admin",
    "team lead",
    "standard with edit access",
]


def assembled_get(path, params=None):
    r = requests.get(f"{ASSEMBLED_BASE}{path}", auth=ASSEMBLED_AUTH, params=params, timeout=30)
    r.raise_for_status()
    return r.json()


def assembled_get_all(path, key):
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


def main():
    if not ASSEMBLED_API_KEY:
        print("Missing ASSEMBLED_API_KEY")
        sys.exit(1)

    print("Fetching roles...")
    roles = assembled_get_all("/roles", "roles")
    role_lookup = {v["id"]: v["name"] for v in roles.values()}
    print(f"  Found {len(roles)} roles")

    print("Fetching people (all pages)...")
    all_people = assembled_get_all("/people", "people")
    print(f"  Found {len(all_people)} people")

    print("")
    print("People who are STAFFABLE but on a non-staffable role:")
    print("-" * 70)

    flagged = []
    for person in all_people.values():
        if person.get("deleted"):
            continue
        if not person.get("staffable"):
            continue

        person_roles = [role_lookup.get(r, "Unknown") for r in person.get("roles", [])]
        role_names_lower = [r.lower() for r in person_roles]

        is_non_staffable_role = any(
            any(nr in r for nr in NON_STAFFABLE_ROLE_NAMES)
            for r in role_names_lower
        )

        if is_non_staffable_role:
            name = f"{person.get('first_name', '')} {person.get('last_name', '')}".strip()
            email = person.get("email", "")
            flagged.append({
                "name": name,
                "email": email,
                "id": person["id"],
                "roles": ", ".join(person_roles),
            })
            print(f"  {name} ({email}) — Role: {', '.join(person_roles)}")

    print("-" * 70)
    print(f"Total flagged: {len(flagged)}")
    print("")
    print("IDs for patching (copy into unstaffable_fix.py):")
    for p in flagged:
        print(f'    "{p["id"]}",  # {p["name"]}')


if __name__ == "__main__":
    main()
