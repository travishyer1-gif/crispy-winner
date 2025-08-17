# Configuration for Microsoft Outlook AI Agent
# IMPORTANT: Replace the placeholder values with your actual credentials
# DO NOT commit this file to version control with real credentials

# Microsoft Azure AD Configuration
TENANT_ID = "your_tenant_id_here"  # Replace with your Azure AD Tenant ID
CLIENT_ID = "your_client_id_here"   # Replace with your Azure AD Application (Client) ID

# User Credentials
USERNAME = "your_email@domain.com"  # Replace with your Microsoft account email
PASSWORD = "your_password_here"      # Replace with your Microsoft account password

# Microsoft Graph API Configuration
GRAPH_API_VERSION = "v1.0"
BASE_URL = "https://graph.microsoft.com"

# Data Retrieval Settings
INBOX_FILTER_KEYWORD = "wisp"  # Keyword to filter inbox emails (modify as needed)
MAX_ITEMS_PER_REQUEST = 100    # Maximum items per API request

# Example of what the values should look like:
# TENANT_ID = "12345678-1234-1234-1234-123456789012"
# CLIENT_ID = "87654321-4321-4321-4321-210987654321"
# USERNAME = "john.doe@company.com"
# PASSWORD = "your_secure_password"
