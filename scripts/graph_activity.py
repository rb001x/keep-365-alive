import os
import requests

TENANT_ID = os.environ.get("MS_TENANT_ID")
CLIENT_ID = os.environ.get("MS_CLIENT_ID")
CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET")

# 1. Acquire Access Token
token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
token_data = {
    "grant_type": "client_credentials",
    "client_id": CLIENT_ID,
    "client_secret": CLIENT_SECRET,
    "scope": "https://graph.microsoft.com/.default"
}

response = requests.post(token_url, data=token_data)
response.raise_for_status()
access_token = response.json().get("access_token")

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}

# 2. Execute Telemetry/Dev API Calls
endpoints = [
    "https://graph.microsoft.com/v1.0/users",
    "https://graph.microsoft.com/v1.0/groups",
    "https://graph.microsoft.com/v1.0/applications",
    "https://graph.microsoft.com/v1.0/organization"
]

print("Starting Microsoft Graph API activity...")
for url in endpoints:
    res = requests.get(url, headers=headers)
    print(f"GET {url} -> Status: {res.status_code}")
    if res.status_code != 200:
        print(f"Warning/Error: {res.text}")

print("Developer activity run completed.")
