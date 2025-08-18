# Configuration Template for Microsoft Outlook AI Agent
# Copy this file to config.py and fill in your actual credentials

# Microsoft Azure AD Configuration
TENANT_ID = "your_tenant_id_here"  # Azure AD Tenant ID
CLIENT_ID = "your_client_id_here"   # Azure AD Application (Client) ID

# User Credentials
USERNAME = "your_email@domain.com"  # Your Microsoft account email
PASSWORD = "your_password_here"      # Your Microsoft account password

# Microsoft Graph API Configuration
GRAPH_API_VERSION = "v1.0"
BASE_URL = "https://graph.microsoft.com"

# Data Retrieval Settings
INBOX_FILTER_KEYWORD = "wisp"  # Keyword to filter inbox emails
MAX_ITEMS_PER_REQUEST = 100    # Maximum items per API request
