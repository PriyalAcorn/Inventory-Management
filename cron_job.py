import time
import requests

# Define constants
CLIENT_ID = "0a95ac99-f2af-4f34-b62b-5c01564e02aa"
CLIENT_SECRET = "******************************" # Add client_secret key
TENANT_ID = "bd262270-d3d4-4e69-8eab-ec0eaec24f0f"
SCOPE = "https://api.businesscentral.dynamics.com/.default"
AUTH_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

def generate_token():
    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
    }

    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials",
        "scope": SCOPE,
    }

    response = requests.post(AUTH_URL, headers=headers, data=data)

    if response.status_code == 200:
        json_response = response.json()
        access_token = json_response.get("access_token")

        if access_token:
            print("Generated Token:", access_token)
            return access_token
        else:
            print("No Access Token Found")
            return None
    else:
        print("Error obtaining access token:", response.text)
        return None

if __name__ == "__main__":
    while True:
        token = generate_token()
        time.sleep(3600)  # Wait for 1 hour before generating a new token
