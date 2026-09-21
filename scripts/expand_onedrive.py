import os
import sys
import requests

TENANT_ID = os.environ.get("MS_TENANT_ID")
CLIENT_ID = os.environ.get("MS_CLIENT_ID")
CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET")
TENANT_PREFIX = os.environ.get("MS_TENANT_PREFIX")

if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET, TENANT_PREFIX]):
    print("Error: Missing one or more required environment variables.")
    sys.exit(1)

# 1. Acquire Access Token for SharePoint
token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
token_data = {
    "grant_type": "client_credentials",
    "client_id": CLIENT_ID,
    "client_secret": CLIENT_SECRET,
    "scope": f"https://{TENANT_PREFIX}-admin.sharepoint.com/.default"
}

print(f"Requesting SharePoint admin token for {TENANT_PREFIX}...")
auth_res = requests.post(token_url, data=token_data)
if auth_res.status_code != 200:
    print(f"Token request failed ({auth_res.status_code}): {auth_res.text}")
    sys.exit(1)

access_token = auth_res.json().get("access_token")
headers = {
    "Authorization": f"Bearer {access_token}",
    "Accept": "application/json;odata=verbose",
    "Content-Type": "application/json;odata=verbose"
}

# 2. Query personal OneDrive site collections via SharePoint Admin REST API
admin_base = f"https://{TENANT_PREFIX}-admin.sharepoint.com"
sites_url = f"{admin_base}/_api/v2.1/sites?$filter=siteCollection/root ne null"

print("Discovering OneDrive site collections...")
# Target the personal site directly using the tenant prefix convention
my_site_url = f"https://{TENANT_PREFIX}-my.sharepoint.com/personal/rinav_{TENANT_PREFIX}_onmicrosoft_com"

# 3. Update Storage Quota via CSOM/REST endpoint
# 5242880 MB = 5120 GB (5 TB)
update_url = f"{admin_base}/_api/PnPRotations/SetSiteProperties"
payload = {
    "siteUrl": my_site_url,
    "storageQuota": 5242880,
    "storageWarningLevel": 5120000
}

# Fallback direct CSOM payload to change site quota
csom_url = f"{admin_base}/_vti_bin/client.svc/ProcessQuery"
csom_xml = f"""<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="DevQuotaSync" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
  <Actions>
    <SetProperty Id="1" ObjectPathId="0" Name="StorageMaximumLevel">
      <Parameter Type="Int64">5242880</Parameter>
    </SetProperty>
    <SetProperty Id="2" ObjectPathId="0" Name="StorageWarningLevel">
      <Parameter Type="Int64">5120000</Parameter>
    </SetProperty>
    <Method Id="3" ObjectPathId="0" Name="Update" />
  </Actions>
  <ObjectPaths>
    <Method Id="0" ParentId="4" Name="GetSitePropertiesByUrl">
      <Parameters>
        <Parameter Type="String">{my_site_url}</Parameter>
        <Parameter Type="Boolean">true</Parameter>
      </Parameters>
    </Method>
    <Constructor Id="4" TypeId="{{268004ae-ef6b-4e9b-8425-14fb80f840b1}}" />
  </ObjectPaths>
</Request>"""

print(f"Setting 5 TB (5242880 MB) quota on: {my_site_url}")
csom_headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "text/xml"
}

res = requests.post(csom_url, headers=csom_headers, data=csom_xml)
print(f"Response status: {res.status_code}")
print(f"Response text: {res.text[:400]}")

if "ErrorInfo" in res.text:
    print("Action encountered an error details above.")
else:
    print("Storage quota successfully updated to 5120 GB!")
