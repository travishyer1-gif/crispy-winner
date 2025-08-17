# Microsoft Outlook AI Agent

An AI-powered agent that integrates with Microsoft Outlook using the Microsoft Graph API to retrieve emails and calendar events.

## Features

- 🔐 **Secure Authentication**: Uses Microsoft Authentication Library (MSAL) for secure OAuth 2.0 authentication
- 📧 **Email Retrieval**: Fetches inbox emails filtered by keyword and sent emails
- 📅 **Calendar Integration**: Retrieves calendar events and appointments
- 📄 **Pagination Handling**: Automatically handles paginated API responses
- 💾 **Data Export**: Saves all retrieved data to structured JSON format
- 🛡️ **Security**: Implements secure credential handling and token management

## Prerequisites

Before using this agent, you need:

1. **Microsoft 365 Account**: A valid Microsoft 365 or Outlook.com account
2. **Azure AD Application**: A registered application in Azure Active Directory
3. **API Permissions**: Appropriate Microsoft Graph API permissions configured

## Setup Instructions

### 1. Azure AD Application Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Fill in the application details:
   - **Name**: Outlook AI Agent
   - **Supported account types**: Accounts in this organizational directory only
   - **Redirect URI**: (Leave blank for this application type)
5. Click **Register**

### 2. Configure API Permissions

1. In your registered application, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Application permissions** or **Delegated permissions**:
   - **Mail.Read** - Read user mail
   - **Calendars.Read** - Read user calendars
   - **User.Read** - Read user profile
5. Click **Grant admin consent**

### 3. Get Application Credentials

1. Note down your **Application (client) ID**
2. Note down your **Directory (tenant) ID**
3. These will be used in the configuration

### 4. Install Dependencies

```bash
pip install -r requirements.txt
```

### 5. Configure Credentials

1. Copy `config_template.py` to `config.py`
2. Fill in your actual credentials:
   ```python
   TENANT_ID = "your_actual_tenant_id"
   CLIENT_ID = "your_actual_client_id"
   USERNAME = "your_email@domain.com"
   PASSWORD = "your_password"
   ```

## Usage

### Basic Usage

```python
from outlook_authenticator import OutlookAuthenticator

# Create authenticator instance
authenticator = OutlookAuthenticator(
    tenant_id="your_tenant_id",
    client_id="your_client_id",
    username="your_email@domain.com",
    password="your_password"
)

# Fetch all data
data = authenticator.fetch_all_data()

# Save to file
authenticator.save_data_to_file(data, "my_outlook_data.json")
```

### Individual Data Retrieval

```python
# Authenticate first
if authenticator.authenticate():
    # Fetch specific data types
    inbox_emails = authenticator.fetch_inbox_emails()
    sent_emails = authenticator.fetch_sent_emails()
    calendar_events = authenticator.fetch_calendar_events()
```

### Command Line Execution

```bash
python outlook_authenticator.py
```

## Data Structure

The retrieved data is saved in the following JSON structure:

```json
{
  "retrieval_timestamp": "2024-01-01T12:00:00",
  "total_items": {
    "inbox_emails": 25,
    "sent_emails": 15,
    "calendar_events": 8
  },
  "inbox_emails": [...],
  "sent_emails": [...],
  "calendar_events": [...]
}
```

## Security Considerations

- **Never commit credentials** to version control
- **Use environment variables** for production deployments
- **Implement proper token caching** for production use
- **Consider using certificate-based authentication** for enterprise scenarios
- **Regularly rotate application secrets**

## Troubleshooting

### Common Issues

1. **Authentication Failed**
   - Verify your Tenant ID and Client ID
   - Check if your account has access to the application
   - Ensure API permissions are granted

2. **Permission Denied**
   - Verify API permissions are configured correctly
   - Check if admin consent is granted
   - Ensure your account has the required permissions

3. **Rate Limiting**
   - The script includes automatic pagination handling
   - Consider implementing delays between requests for large datasets

### Error Messages

- `"No access token available"`: Run `authenticate()` first
- `"Authentication failed"`: Check credentials and permissions
- `"API request failed"`: Verify network connectivity and API endpoints

## API Endpoints Used

- **Authentication**: `https://login.microsoftonline.com/{tenant_id}`
- **Inbox Emails**: `/me/messages` with subject filter
- **Sent Emails**: `/me/mailFolders('sentitems')/messages`
- **Calendar Events**: `/me/events`

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For issues and questions:
1. Check the troubleshooting section
2. Review Microsoft Graph API documentation
3. Open an issue in the repository

## Project Status

### ✅ Step 1: Complete - Microsoft Graph API Authentication and Data Retrieval

**Status**: ✅ **IMPLEMENTATION COMPLETE** - Awaiting Azure AD Permissions

This step has been fully implemented and includes:

- **MSAL Authentication**: Secure OAuth 2.0 authentication using Microsoft Authentication Library
- **Client Secret Authentication**: Uses Azure AD application permissions for secure access
- **Email Retrieval**: Fetches inbox emails filtered by 'wisp' keyword and all sent emails
- **Calendar Integration**: Retrieves all calendar events and appointments
- **Pagination Handling**: Automatically handles paginated API responses using @odata.nextLink
- **Data Export**: Saves all retrieved data to structured JSON format (`outlook_data.json`)
- **Error Handling**: Comprehensive error handling and user feedback
- **Configuration Management**: Secure credential handling through config.py

**Files Created/Modified**:
- `outlook_authenticator.py` - Main implementation class (255 lines)
- `config_template.py` - Configuration template
- `config.py` - User configuration (create from template)
- `test_step1.py` - Test script to verify functionality
- `requirements.txt` - Required Python packages

**Current Status**:
- ✅ **Authentication**: Working perfectly with client secret
- ❌ **API Access**: Waiting for Azure AD permissions (Mail.Read.All, Calendars.Read.All)
- 📊 **Implementation**: 100% complete and ready for testing

**Testing**:
```bash
# Run the test script to verify Step 1 functionality
python test_step1.py

# Or run the main authenticator directly
python outlook_authenticator.py
```

**Next Action Required**:
1. **Configure Azure AD permissions** in Azure Portal:
   - Mail.Read.All
   - Calendars.Read.All
   - User.Read.All
2. **Grant admin consent** for the permissions
3. **Test the complete implementation**

### 🔄 Next Steps

Future steps in the AI Agent project will include:
- **Step 2**: Natural language processing of email content
- **Step 3**: Intelligent email categorization and prioritization
- **Step 4**: Automated response generation
- **Step 5**: Calendar optimization and scheduling assistance
- **Step 6**: Integration with other AI services
