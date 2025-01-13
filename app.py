import json, requests
from msal import PublicClientApplication

# Constants
CLIENT_ID = '1c8d3638-0314-474a-9dab-f8740644912a'
TENANT_ID = '241acc2f-5659-4390-9160-fcde536b90fe'
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Calendars.ReadWrite"]
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0/me/events'

# Initialize the Public Client Application
app = PublicClientApplication(client_id=CLIENT_ID, authority=AUTHORITY)

# Acquire token interactively
accounts = app.get_accounts()

for account in app.get_accounts():
    app.remove_account(account)

if accounts:
    # Optionally handle existing accounts for silent token acquisition
    token_response = app.acquire_token_silent(SCOPES, account=accounts[0])
else:
    token_response = app.acquire_token_interactive(SCOPES)

# Check if token acquisition was successful
if "access_token" in token_response:
    access_token = token_response['access_token']
    print("Access token acquired successfully.")

    # Decode the token to check scopes
    import jwt
    decoded_token = jwt.decode(access_token, options={"verify_signature": False})
    print("Token scopes:", decoded_token.get('scp'))

    # Event details
    event = {
        "subject": "Team Meeting",
        "body": {"contentType": "HTML", "content": "Discuss project updates"},
        "start": {"dateTime": "2025-01-10T14:00:00", "timeZone": "Pacific Standard Time"},
        "end": {"dateTime": "2025-01-10T15:00:00", "timeZone": "Pacific Standard Time"},
        "attendees": [
            {
                "emailAddress": {"address": "attendee@example.com", "name": "Attendee Name"},
                "type": "required"
            }
        ]
    }

    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.post(GRAPH_API_ENDPOINT, headers=headers, data=json.dumps(event))

    if response.status_code == 201:
        print("Appointment booked successfully.")
    else:
        print(f"Failed to book appointment: {response.status_code}")
        print(response.json())
else:
    print("Failed to acquire access token.")
    print(token_response.get('error'))
    print(token_response.get('error_description'))