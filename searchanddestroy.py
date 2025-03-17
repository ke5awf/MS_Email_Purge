import requests
import json
import time
from datetime import datetime, timedelta
from msal import ConfidentialClientApplication
from requests.exceptions import HTTPError
from urllib.parse import quote



# Azure AD Configuration
CLIENT_ID = 'CLIENT_ID'
CLIENT_SECRET = 'CLIENT SECRET'
TENANT_ID = 'TENANT_ID'
# Microsoft Graph API Endpoint
GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0'

# Retry logic for 503 errors
def make_request_with_retries(url, headers, max_retries=5, delay=5):
    for attempt in range(max_retries):
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()  # Raise HTTPError for bad responses
            return response
        except HTTPError as e:
            if response.status_code == 503:
                print(f"503 Error: Retrying in {delay} seconds... (Attempt {attempt + 1}/{max_retries})")
                time.sleep(delay)
                delay *= 2  # Exponential backoff
            else:
                raise e
    raise Exception("Failed to retrieve data after multiple attempts.")

# Authentication
def get_access_token():
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=f'https://login.microsoftonline.com/{TENANT_ID}',
        client_credential=CLIENT_SECRET
    )
    token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

    if 'access_token' in token_response:
        return token_response['access_token']
    else:
        raise Exception(f"Failed to retrieve token: {token_response.get('error_description', 'Unknown error')}")

# Pagination support for large result sets
def get_all_emails(url, headers):
    emails = []
    while url:
        response = make_request_with_retries(url, headers)
        data = response.json()
        emails.extend(data.get('value', []))
        url = data.get('@odata.nextLink', None)  # Continue pagination if more data exists
    return emails

# Search and delete email
def search_and_delete_email(sender, subject, date_option, mailbox=None):
    headers = {"Authorization": f"Bearer {get_access_token()}"}

    # Determine date filter based on user choice
    date_filter = ""
    if date_option == 'exact':
        date = input("Enter exact date (YYYY-MM-DD): ")
        date_filter = (
            f"and receivedDateTime ge {date}T00:00:00Z "
            f"and receivedDateTime le {date}T23:59:59Z"
        )
    elif date_option == 'last7':
        seven_days_ago = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
        date_filter = f"and receivedDateTime ge {seven_days_ago}T00:00:00Z"

    # Encode subject to avoid issues with special characters
    subject_filter = f"and subject eq '{quote(subject)}'" if subject else ""

    # Correct filter syntax
    query = (
        f"from/emailAddress/address eq '{sender}' "
        f"{subject_filter} "
        f"{date_filter}"
    ).strip()

    if mailbox:
        # Search in specific mailbox
        url = f"{GRAPH_ENDPOINT}/users/{mailbox}/messages?$filter={query}"
    else:
        # Search across all mailboxes in the tenant (requires special permissions)
        url = f"{GRAPH_ENDPOINT}/security/emails?$filter={query}"

    emails = get_all_emails(url, headers)

    if not emails:
        print("No matching emails found.")
        return

    for email in emails:
        email_id = email['id']
        email_subject = email['subject']
        print(f"Found email: {email_subject}")

        confirm = input(f"Do you want to delete this email? (yes/no): ")
        if confirm.lower() == 'yes':
            delete_url = f"{GRAPH_ENDPOINT}/users/{mailbox}/messages/{email_id}" if mailbox else f"{GRAPH_ENDPOINT}/security/emails/{email_id}"
            delete_response = requests.delete(delete_url, headers=headers)
            if delete_response.status_code == 204:
                print(f"Successfully deleted: {email_subject}")
            else:
                print(f"Failed to delete: {email_subject}")
        else:
            print("Skipping deletion.")

if __name__ == "__main__":
    sender = input("Enter sender email: ")
    subject = input("Enter email subject (optional): ").strip()
    date_option = input("Search by exact date, within the last 7 days, or leave blank for no date filter? (exact/last7/none): ").strip().lower()
    mailbox = input("Enter mailbox (leave blank for all mailboxes): ").strip() or None

    search_and_delete_email(sender, subject, date_option, mailbox)

