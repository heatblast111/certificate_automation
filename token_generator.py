from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

flow = InstalledAppFlow.from_client_secrets_file("creds.json", SCOPES)
creds = flow.run_local_server(port=0)

# Save the token
with open("token.json", "w") as token_file:
    token_file.write(creds.to_json())

print("✅ token.json generated successfully")
