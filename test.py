from cleanser import GraphAuth, GRAPH_BASE
import requests
import json
import base64

auth = GraphAuth()
token = auth.get_token()
headers = {"Authorization": f"Bearer {token}"}

# Decode the token to see what account/tenant we're hitting
print("=== Token Info ===")
try:
    payload = token.split(".")[1]
    # Fix padding
    payload += "=" * (4 - len(payload) % 4)
    decoded = json.loads(base64.b64decode(payload))
    print(f"  Tenant ID: {decoded.get('tid', '?')}")
    print(f"  User:      {decoded.get('upn', decoded.get('unique_name', decoded.get('preferred_username', '?')))}")
    print(f"  App ID:    {decoded.get('appid', decoded.get('azp', '?'))}")
    print(f"  Audience:  {decoded.get('aud', '?')}")
except Exception as e:
    print(f"  Could not decode token: {e}")
print()

# Full /me profile
print("=== Full /me Profile ===")
r = requests.get(f"{GRAPH_BASE}/me", headers=headers)
me = r.json()
for key in ["displayName", "mail", "userPrincipalName", "id", "jobTitle", "officeLocation"]:
    print(f"  {key}: {me.get(key)}")
print()

# Check mailbox settings
print("=== Mailbox Settings ===")
r = requests.get(f"{GRAPH_BASE}/me/mailboxSettings", headers=headers)
if r.status_code == 200:
    settings = r.json()
    print(f"  Timezone: {settings.get('timeZone')}")
    print(f"  Date format: {settings.get('dateFormat')}")
    auto = settings.get("automaticRepliesSetting", {})
    print(f"  Auto-replies status: {auto.get('status')}")
else:
    print(f"  Error {r.status_code}: {r.text[:200]}")
print()

# Count messages across all folders
print("=== Message Counts ===")
r = requests.get(
    f"{GRAPH_BASE}/me/messages",
    headers=headers,
    params={"$count": "true", "$top": 1},
)
data = r.json()
print(f"  Total messages via /me/messages: {data.get('@odata.count', len(data.get('value', [])))}")
print()

# List folders again with full detail
print("=== All Folders ===")
r = requests.get(
    f"{GRAPH_BASE}/me/mailFolders",
    headers=headers,
    params={"$top": 100, "includeHiddenFolders": "true"}
)
folders = r.json().get("value", [])
total_across_folders = 0
for f in folders:
    name = f.get("displayName", "?")
    total = f.get("totalItemCount", 0)
    total_across_folders += total
    if total > 0 or f.get("childFolderCount", 0) > 0:
        print(f"  {name:<30} total: {total:>8}  children: {f.get('childFolderCount', 0)}")
print(f"\n  Sum across all visible folders: {total_across_folders}")