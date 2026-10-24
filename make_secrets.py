"""
Builds .streamlit/secrets.toml for Glovebox.
Run from the glovebox folder:  python make_secrets.py
It reads your Google key (.json) and asks for the rest.
"""
import glob
import json
import os

keys = [f for f in glob.glob("*.json") if "service_account" in open(f).read()]
if not keys:
    raise SystemExit("No Google key (.json) found in this folder. Move it here and run again.")
key_file = keys[0]
key = json.load(open(key_file))
print(f"Using Google key: {key_file}\n")


def ask(prompt, required=True):
    while True:
        v = input(prompt).strip().strip('"').strip("'")
        if v or not required:
            return v
        print("  This one is needed.")


sheet_id = ask("Sheet ID (from the sheet's URL): ")
if "/d/" in sheet_id:  # pasted the whole URL
    sheet_id = sheet_id.split("/d/")[1].split("/")[0]
password = ask("Household password for the app: ")
names = ask("Your names, separated by commas (e.g. Justin, Sarah): ")
drivers = [n.strip() for n in names.split(",") if n.strip()]
url = ask("Receipts Web app URL (ends in /exec, Enter to skip): ", required=False)
url = url.replace("/u/0/", "/").replace("/u/1/", "/").replace("/u/2/", "/").replace("/u/3/", "/")
token = ask("Secret phrase from the script's TOKEN line: ", required=bool(url)) if url else ""


def q(s):
    return json.dumps(s)  # safe TOML string with escaped newlines


lines = [
    f"app_password = {q(password)}",
    f"drivers = {json.dumps(drivers)}",
    f"sheet_id = {q(sheet_id)}",
    f"receipts_url = {q(url)}",
    f"receipts_token = {q(token)}",
    "",
    "[gcp_service_account]",
] + [f"{k} = {q(v)}" for k, v in key.items()]

os.makedirs(".streamlit", exist_ok=True)
with open(".streamlit/secrets.toml", "w") as f:
    f.write("\n".join(lines) + "\n")

print("\n✅ Wrote .streamlit/secrets.toml")
print("Next: streamlit run app.py")
