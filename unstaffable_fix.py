import os
import sys
import requests

ASSEMBLED_API_KEY = os.environ.get("ASSEMBLED_API_KEY")
ASSEMBLED_BASE = "https://api.assembledhq.com/v0"
ASSEMBLED_AUTH = (ASSEMBLED_API_KEY, "")

# 27 people to set as unstaffable
UNSTAFFABLE = [
    ("dc84c2a3-2a52-4e7a-8cf9-de266359f1af", "Charley Rooney"),
    ("158fd839-3995-4dbb-80fc-8d4c9720099b", "Monika Kelemen"),
    ("d4a36801-37a2-44d2-927b-e5d58b91db32", "Ingrid Kocsardi"),
    ("23d9deb5-58e7-4e17-9aa3-82ea723b3d47", "Ildiko Nagy"),
    ("c019dd60-6ce4-4284-bc53-cdcd0b66a511", "Reka Palinkas"),
    ("3c893772-4f6a-498e-b104-70affd916b80", "Eszter Vasvari"),
    ("1232ffcd-f3de-43ad-9ad0-2103a44ac249", "Adam Szoke"),
    ("a748fb9e-6fb8-464a-9eee-df08d23e0253", "Eszter Brhlik"),
    ("3d45ebe5-ec09-47fd-8e96-aa0c2e6aea28", "Bori Bardos"),
    ("ee058635-5816-45d4-9661-1c1e2d7aac44", "Ferenc Hegyvari"),
    ("1d5d5ac6-5bac-4809-8443-9a50e70fd539", "Ivett Kalan"),
    ("3f1b852f-0061-4e43-976d-da5543637fa8", "Gabriela da Silveira"),
    ("93044f16-82c9-4c3e-a7c2-0ccb77c304db", "Kirill Penchukov"),
    ("b5b4020b-d4b4-4d0f-9d30-1a916b7ae450", "Anji Sardeson"),
    ("015f1b70-1d1a-4fb8-8f32-5f980616df61", "Kunal Varia"),
    ("30e09a9f-2e14-4bbd-8f69-a72925e3e6f8", "Jamie Perry"),
    ("3d6b9643-e03b-44dc-8af7-626e4c033eed", "Laura Maxfield"),
    ("1526ac6b-d7d9-4a67-a4b0-094f4a3ca766", "Ella Raffan"),
    ("12f7ae93-0f41-44f5-906f-e5df69497ffe", "Mate Tomor"),
    ("10de1207-1f1d-46b7-8572-fbd2906b2333", "Lauren Gatt"),
    ("66b4913d-e96d-434a-8b2f-0732840b93ee", "Charles Smallwood"),
    ("16d96409-444d-467b-b935-48fec352cb4b", "Khalid Abouargub"),
    ("eaee3777-092d-4c94-ac33-a846c3e9f8e4", "Hanna Suhajda"),
    ("c37dbd11-2f21-4147-95d1-c8e5c3594756", "Vivien Stroban"),
    ("b7719876-ce72-4ac6-a334-2775df74af17", "Anna Joyner"),
    ("241acd7b-f1f2-49e3-9677-c05b11f01fef", "Akos Nagy"),
    ("93354367-5d0d-4eee-abe9-dc80e72da41d", "Kristof Snee"),
]


def assembled_patch(path, payload):
    r = requests.patch(
        f"{ASSEMBLED_BASE}{path}",
        auth=ASSEMBLED_AUTH,
        json=payload,
        timeout=30
    )
    r.raise_for_status()
    return r.json()


def main():
    if not ASSEMBLED_API_KEY:
        print("Missing ASSEMBLED_API_KEY")
        sys.exit(1)

    print("=============================================")
    print("  Unstaffable Fix — setting staffable: false")
    print(f"  {len(UNSTAFFABLE)} people to update")
    print("=============================================")

    updated = 0
    failed = 0

    for person_id, name in UNSTAFFABLE:
        try:
            assembled_patch(f"/people/{person_id}", {"staffable": False})
            print(f"  OK  {name}")
            updated += 1
        except Exception as e:
            print(f"  ERR {name} — {e}")
            failed += 1

    print("")
    print(f"-- Done ----------------------------------------")
    print(f"  Updated: {updated}")
    print(f"  Failed:  {failed}")
    print(f"------------------------------------------------")


if __name__ == "__main__":
    main()
