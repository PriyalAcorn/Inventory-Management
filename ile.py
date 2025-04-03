import requests
import pyodbc
import json

# Fetch inventory data
url = "https://api.businesscentral.dynamics.com/v2.0/bd262270-d3d4-4e69-8eab-ec0eaec24f0f/Sandbox1911/ODataV4/Company('Jambo%20Supplies%20Ltd.')/Item_Ledger_Entries_Excel"
headers = {"Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IkNOdjBPSTNSd3FsSEZFVm5hb01Bc2hDSDJYRSIsImtpZCI6IkNOdjBPSTNSd3FsSEZFVm5hb01Bc2hDSDJYRSJ9.eyJhdWQiOiJodHRwczovL2FwaS5idXNpbmVzc2NlbnRyYWwuZHluYW1pY3MuY29tIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvYmQyNjIyNzAtZDNkNC00ZTY5LThlYWItZWMwZWFlYzI0ZjBmLyIsImlhdCI6MTc0MzY2NjQ4NywibmJmIjoxNzQzNjY2NDg3LCJleHAiOjE3NDM2NzAzODcsImFpbyI6ImsyUmdZSGh4MXROOXk2cFFFLy9xV3JlbzYzdHRBUT09IiwiYXBwaWQiOiIwYTk1YWM5OS1mMmFmLTRmMzQtYjYyYi01YzAxNTY0ZTAyYWEiLCJhcHBpZGFjciI6IjEiLCJpZHAiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9iZDI2MjI3MC1kM2Q0LTRlNjktOGVhYi1lYzBlYWVjMjRmMGYvIiwiaWR0eXAiOiJhcHAiLCJvaWQiOiJlOGEwNTZlNi00NDBlLTQwODYtOTliMC1kZWYwOTQ1NTRlYTciLCJyaCI6IjEuQVNzQWNDSW12ZFRUYVU2T3Etd09yc0pQRHozdmJabHNzMU5CaGdlbV9Ud0J1SjhyQUFBckFBLiIsInJvbGVzIjpbIkF1dG9tYXRpb24uUmVhZFdyaXRlLkFsbCIsImFwcF9hY2Nlc3MiLCJBZG1pbkNlbnRlci5SZWFkV3JpdGUuQWxsIiwiQVBJLlJlYWRXcml0ZS5BbGwiXSwic3ViIjoiZThhMDU2ZTYtNDQwZS00MDg2LTk5YjAtZGVmMDk0NTU0ZWE3IiwidGlkIjoiYmQyNjIyNzAtZDNkNC00ZTY5LThlYWItZWMwZWFlYzI0ZjBmIiwidXRpIjoiekVQMktOU0N4RVdwRHZZYjFoTVVBQSIsInZlciI6IjEuMCIsInhtc19pZHJlbCI6IjggNyJ9.daikFL3E_Wrq2w-0ojkJJ2LJ2rjvTKQe9Mh3lIP8femhJM8_ak-HoxmupJBaKxJwLx0XAvdtH__7Kcep3IaZ9ldbdpDjRu_0e0svC9IqkqqhG8v07Tz3v2m3UBxYpCiAv3vx32n2L744pwm2jdqnhCZ5duHlHtC9SE-4r9rpTG_DM38qvexJLI5wwPKewTG94dOR6kDo4cHmUlA9LUGSstHGmAVlCi9bQJC0QYzXTRqsJKjO-eq3Wu4ZEDssMV2RHCLiEttsd3bB58844DSe6YpZh9FP7XqOQYVlpP3JqA_C3pRVRW6sFIXV5WVcJRhZtnD1ffHuVix19xjjnnauhA"}
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