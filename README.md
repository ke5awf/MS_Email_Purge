# MS_Email_Purge
Microsoft Graph Email Search and Delete


This tool is to be able to quickly serach and delete emails from your Exchange Online. 

Instructions for Setup

Dependencies

    pip install msal requests

Register an Azure AD App

    Go to Azure Portal.
    Register a new app.
    Under API permissions, add:
        Mail.ReadWrite (for mailbox-specific search and deletion)
        EWS.AccessAsUser.All (for tenant-wide search and deletion)
    Grant admin consent for the permissions.

Create and Assign a Security Role

    For tenant-wide searches, assign the EWS Impersonation role or Global Admin.

Environment Configuration

    Replace:
        CLIENT_ID with your Azure AD app's Client ID.
        CLIENT_SECRET with your Azure AD app's Secret.
        TENANT_ID with your organization's Tenant ID.

Running the Script Run the script
