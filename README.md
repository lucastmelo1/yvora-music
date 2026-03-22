# Yvora Music - single app

Single codebase and single link.

Secrets supported:

Option A
[google]
sheet_id = "..."
service_account_json = '''{...json...}'''

Option B
[google]
sheet_id = "..."

[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = '''-----BEGIN PRIVATE KEY-----
...
-----END PRIVATE KEY-----'''
client_email = "..."
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."

Google Sheets tabs:
users, sessions, chapters, live, gemini_history
