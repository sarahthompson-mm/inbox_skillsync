import os, requests

ASSEMBLED_API_KEY = os.environ.get("ASSEMBLED_API_KEY")
ASSEMBLED_BASE = "https://api.assembledhq.com/v0"
ASSEMBLED_AUTH = (ASSEMBLED_API_KEY, "")

results = {}
offset = 0
while True:
    r = requests.get(f"{ASSEMBLED_BASE}/people", auth=ASSEMBLED_AUTH, 
        params={"limit": 100, "offset": offset}, timeout=30)
    data = r.json()
    page = data.get("people", {})
    results.update(page)
    total = data.get("total", 0)
    offset += 100
    if offset >= total:
        break

for person in results.values():
    name = f"{person.get('first_name','')} {person.get('last_name','')}".strip()
    if name in ["Anna Joyner", "Akos Nagy"]:
        print(f"{name}: staffable={person.get('staffable')} agent_role={person.get('agent_role')}")
