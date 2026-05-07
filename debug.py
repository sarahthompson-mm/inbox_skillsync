import os, requests
ASSEMBLED_API_KEY = os.environ.get("ASSEMBLED_API_KEY")
r = requests.get("https://api.assembledhq.com/v0/people", auth=(ASSEMBLED_API_KEY, ""), timeout=30)
data = r.json()
people = data.get("people", {})
print(f"Total people returned by API: {len(people)}")
print(f"Response keys: {list(data.keys())}")
# Print first 3 raw people so we can see the full structure
for i, (k, p) in enumerate(people.items()):
    if i >= 3:
        break
    print(f"\nKey: {k}")
    print(f"  name: {p.get('first_name')} {p.get('last_name')}")
    print(f"  platforms: {p.get('platforms')}")
    print(f"  deleted: {p.get('deleted')}")
