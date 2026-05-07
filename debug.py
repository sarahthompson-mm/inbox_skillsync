import os, requests
ASSEMBLED_API_KEY = os.environ.get("ASSEMBLED_API_KEY")
r = requests.get("https://api.assembledhq.com/v0/people", auth=(ASSEMBLED_API_KEY, ""), timeout=30)
people = r.json().get("people", {}).values()
for p in people:
    name = f"{p.get('first_name','')} {p.get('last_name','')}".strip()
    if "anna" in name.lower() or "fekete" in name.lower():
        print(f"Found: {name}")
        print(f"  id: {p.get('id')}")
        print(f"  platforms: {p.get('platforms')}")
        print(f"  deleted: {p.get('deleted')}")
