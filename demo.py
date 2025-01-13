from flask import Flask, redirect, request, url_for, session
from msal import ConfidentialClientApplication
import requests

app = Flask(__name__)
app.secret_key = "your_secret_key"  # Replace with a secure random key

CLIENT_ID = "d9b66f51-647a-4d70-82db-f864d35503c1"
TENANT_ID = "241acc2f-5659-4390-9160-fcde536b90fe"
CLIENT_SECRET = "a0o8Q~Xw3yOMjNiY2lrPWDG7YCPK3QbciOD5KbOC"

AUTHORITY = f"https://login.microsoftonline.com/common"
GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"

REDIRECT_URI = "http://localhost:5000/getAToken"
SCOPES = ["Calendars.ReadWrite", "User.Read", "Mail.ReadWrite", "Mail.Send"]

msal_app = ConfidentialClientApplication(
    CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
)

@app.route("/")
def index():
    if "access_token" in session:
        return "You are logged in. <a href='/create_event'>Create Event</a>"
    else:
        return "<a href='/login'>Log in with Microsoft</a>"

@app.route("/login")
def login():
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI)
    return redirect(auth_url)

@app.route("/getAToken")
def authorized():
    code = request.args.get("code")
    token_response = msal_app.acquire_token_by_authorization_code(
        code, scopes=SCOPES, redirect_uri=REDIRECT_URI
    )
    if "access_token" in token_response:
        session["access_token"] = token_response["access_token"]
        return redirect(url_for("index"))
    else:
        return f"Login failed: {token_response}"

@app.route("/create_event")
def create_event():
    if "access_token" not in session:
        return redirect(url_for("login"))

    access_token = session["access_token"]
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    event_data = {
        "subject": "Test Event",
        "start": {"dateTime": "2025-01-13T10:00:00", "timeZone": "UTC"},
        "end": {"dateTime": "2025-01-13T11:00:00", "timeZone": "UTC"},
        "location": {"displayName": "Virtual Meeting"},
        "attendees": [
            {
                "emailAddress": {"address": "pallavi@articence.com", "name": "Pallavi"},
                "type": "required"
            }
        ]
    }

    response = requests.post(f"{GRAPH_ENDPOINT}/me/events", headers=headers, json=event_data)

    if response.status_code == 201:
        return "Event created successfully!"
    else:
        return f"Failed to create event: {response.status_code}, {response.text}"

if __name__ == "__main__":
    app.run(port=5000, debug=True)
