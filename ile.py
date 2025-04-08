import requests
import pyodbc
import json
from cron_job import generate_token  # Import the function from cron_job.py

# First commit to test

# Fetch inventory data
url = "https://api.businesscentral.dynamics.com/v2.0/bd262270-d3d4-4e69-8eab-ec0eaec24f0f/Sandbox1911/ODataV4/Company('Jambo%20Supplies%20Ltd.')/Item_Ledger_Entries_Excel"
token = generate_token()  # Get the latest token
headers = {"Authorization": f"Bearer {token}"}
response = requests.get(url, headers=headers)

if response.status_code == 200:
    try:
        full_data = response.json()
        if "value" in full_data:
            data = full_data["value"]
        else:
            print("❌ Error: 'value' key not found in API response!")
            exit()

        if not data:
            print("⚠️ No inventory data found!")
            exit()

        # Debug: Print first record to identify correct field names
        print("🔍 First record from API:", json.dumps(data[0], indent=4))

    except json.JSONDecodeError:
        print("❌ Error: API response is not valid JSON")
        exit()

    # Connect to SQL Server
    conn = pyodbc.connect("DRIVER={SQL Server};SERVER=bcdwsqlserver.database.windows.net;DATABASE=BCDWsqldatabase;UID=bcsqlsvr;PWD=Ac0rn$ql188")
    cursor = conn.cursor()

    # Identify correct field names
    for item in data:
        print(f"📦 Processing Item: {item}")  # Debugging each record

        # Correct field names based on API response
        Item_No = item.get("Item_No") or item.get("ItemNumber") or item.get("No.") or "UNKNOWN"
        Description = item.get("Description", "No Description")
        Quantity = item.get("Quantity", 0)
        Location_Code = item.get("Location_Code") or item.get("Location") or "Unknown"

        if Item_No == "UNKNOWN":
            print(f"⚠️ Skipping invalid record: {item}")
            continue

        # Insert into SQL Server
        cursor.execute("""
            INSERT INTO InventoryTable (Item_No, Description, Quantity, Location_Code)
            VALUES (?, ?, ?, ?)
        """, Item_No, Description, Quantity, Location_Code)

    conn.commit()
    print("✅ Inventory data updated successfully!")
    cursor.close()
    conn.close()

else:
    print(f"❌ Failed to fetch data! Status Code: {response.status_code}")
    print("Response Text:", response.text)