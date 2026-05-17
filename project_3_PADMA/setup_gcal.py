"""First-time Google Calendar authorization.

Steps to get credentials.json:
  1. Go to https://console.cloud.google.com/
  2. Create (or select) a project
  3. Enable "Google Calendar API"
  4. Go to APIs & Services > Credentials > Create Credentials > OAuth 2.0 Client ID
  5. Application type: Desktop App
  6. Download the JSON → save as:  persistent_mem/credentials.json
  7. Run this script:  uv run python setup_gcal.py
     A browser window opens — sign in and grant calendar access.
     The token is saved to persistent_mem/token.json for future use.
"""
from pathlib import Path

CREDS_PATH = Path("persistent_mem/credentials.json")
TOKEN_PATH  = Path("persistent_mem/token.json")
SCOPES      = ["https://www.googleapis.com/auth/calendar"]


def main() -> None:
    if not CREDS_PATH.exists():
        print(f"\nERROR: {CREDS_PATH} not found.\n")
        print(__doc__)
        return

    try:
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError:
        print("Run:  uv add google-auth-oauthlib google-api-python-client google-auth-httplib2")
        return

    flow = InstalledAppFlow.from_client_secrets_file(str(CREDS_PATH), SCOPES)
    creds = flow.run_local_server(port=0)
    TOKEN_PATH.parent.mkdir(parents=True, exist_ok=True)
    TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")
    print(f"\nAuthorization successful!  Token saved -> {TOKEN_PATH}")
    print("The create_calendar_event tool will now use your Google Calendar.")


if __name__ == "__main__":
    main()
