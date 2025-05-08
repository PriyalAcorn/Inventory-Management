import pandas as pd
import pyodbc
import os
import base64
import msal
import requests
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# ------------------ SQL Query + Excel Export ------------------

# Connection (using SQL Server Authentication)
conn_str = (
    f"Driver={{SQL Server}};"
    f"Server={os.getenv('SQL_SERVER')};"
    f"Database={os.getenv('SQL_DATABASE')};"
    f"UID={os.getenv('SQL_USERNAME')};"
    f"PWD={os.getenv('SQL_PASSWORD')};"
)
conn = pyodbc.connect(conn_str)

query = """
SELECT 
    s.[seller-sku], 
    s.[UNITS],
    a.[Reference Variant Code],
    a.[Item No_], 
    a.[Description], 
    a.[PK-Size], 
    a.[Single Item Code],
    c.[Remaining_QTY],
    CASE 
        WHEN c.[Remaining_QTY] > 1000 THEN ROUND(c.[Remaining_QTY] / CAST(a.[PK-Size] AS FLOAT), 0)
        ELSE NULL  -- or use 0 if you'd prefer that
    END AS Quantity
FROM 
    [dbo].[SellerSKU] s
INNER JOIN (
    SELECT 
        a.[No_] AS [Item No_],
        a.[Description],
        a.[PK-Size],
        a.[Single Item Code],
        b.[Reference Variant Code]
    FROM 
        [dbo].[Acorn Solutions Ltd_ GB$Item$437dbf0e-84ff-417a-965d-ed2bb9650972] a
    INNER JOIN 
        [dbo].[Acorn Solutions Ltd_ GB$Item Variant$437dbf0e-84ff-417a-965d-ed2bb9650972] b
        ON a.[No_] = b.[Item No_]
) a ON s.[seller-sku] = a.[Reference Variant Code]
INNER JOIN 
    [dbo].[Upload_Tbl_InventoryAPIAuto] c
    ON c.Item_No = a.[Single Item Code]
WHERE 
    s.UNITS != 0
"""

df = pd.read_sql_query(query, conn)
print(df.head())

# Save to Excel
save_path = r'\\172.30.36.151\Amazon Data\Acorn Solution Limited\Inventory\inventory_output.xlsx'
df.to_excel(save_path, index=False)
print(f"Excel file saved successfully at {save_path}")

# ------------------ Graph API Email Logic ------------------

# Load Azure App credentials from .env file
client_id = os.getenv('CLIENT_ID')
tenant_id = os.getenv('TENANT_ID')
client_secret = os.getenv('CLIENT_SECRET')
authority = f'https://login.microsoftonline.com/{tenant_id}'
scopes = ['https://graph.microsoft.com/.default']

# Shared mailbox and recipient from .env file
shared_mailbox = os.getenv('SHARED_MAILBOX')
recipient_email = os.getenv('RECIPIENT_EMAIL')

# MSAL authentication (using ConfidentialClientApplication with client secret)
app = msal.ConfidentialClientApplication(
    client_id,
    authority=authority,
    client_credential=client_secret  # Using client secret for confidential app
)

# Acquire token using client credentials flow
result = app.acquire_token_for_client(scopes=scopes)

if 'access_token' not in result:
    raise Exception("Failed to obtain token:\n" + str(result))

# Read attachment
with open(save_path, 'rb') as f:
    file_bytes = f.read()
file_name = os.path.basename(save_path)

# Prepare Graph API attachment
attachment = {
    "@odata.type": "#microsoft.graph.fileAttachment",
    "name": file_name,
    "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "contentBytes": base64.b64encode(file_bytes).decode('utf-8')
}

# Compose email
email_msg = {
    "message": {
        "subject": "Inventory Export",
        "body": {
            "contentType": "Text",
            "content": "Please find attached the latest inventory export."
        },
        "toRecipients": [
            {
                "emailAddress": {
                    "address": recipient_email
                }
            }
        ],
        "ccRecipients": [
            {
                "emailAddress": {
                    "address": "tanvi.laddha@acornuniversalconsultancy.com"
                }
            },
            {
                "emailAddress": {
                    "address": "aniruddh.toke@acornuniversalconsultancy.com"
                }
            }
        ],
        "from": {
            "emailAddress": {
                "address": shared_mailbox
            }
        },
        "attachments": [attachment]
    },
    "saveToSentItems": "true"
}

# Send email from authenticated app, *as* the shared mailbox
headers = {
    'Authorization': 'Bearer ' + result['access_token'],
    'Content-Type': 'application/json'
}

# Use the shared mailbox address in the URL
send_url = f"https://graph.microsoft.com/v1.0/users/{shared_mailbox}/sendMail"

response = requests.post(send_url, headers=headers, json=email_msg)

# Result
if response.status_code == 202:
    print("✅ Mail sent successfully!")
else:
    print(f"❌ Failed to send email: {response.status_code}")
    print(response.text)