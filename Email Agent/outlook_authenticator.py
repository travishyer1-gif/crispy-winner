import msal
import requests
import json
import os
from datetime import datetime
from typing import List, Dict, Any

class OutlookAuthenticator:
    def __init__(self, tenant_id: str, client_id: str, client_secret: str):
        """
        Initialize the Outlook Authenticator with Microsoft Graph API credentials.
        
        Args:
            tenant_id (str): Azure AD Tenant ID
            client_id (str): Azure AD Application (Client) ID
            client_secret (str): Azure AD Client Secret
        """
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"
        self.scope = ["https://graph.microsoft.com/.default"]
        self.access_token = None
        
    def authenticate(self) -> bool:
        """
        Authenticate using MSAL and acquire access token via client credentials flow.
        
        Returns:
            bool: True if authentication successful, False otherwise
        """
        try:
            # Create MSAL Confidential Client Application
            app = msal.ConfidentialClientApplication(
                client_id=self.client_id,
                client_credential=self.client_secret,
                authority=self.authority
            )
            
            # Acquire token using client credentials flow
            result = app.acquire_token_for_client(scopes=self.scope)
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                print("✅ Authentication successful!")
                return True
            else:
                print(f"❌ Authentication failed: {result.get('error_description', 'Unknown error')}")
                return False
                
        except Exception as e:
            print(f"❌ Authentication error: {str(e)}")
            return False
    
    def _make_graph_request(self, endpoint: str, params: Dict[str, Any] = None) -> List[Dict[str, Any]]:
        """
        Make a request to Microsoft Graph API with pagination handling.
        
        Args:
            endpoint (str): Graph API endpoint
            params (dict): Query parameters
            
        Returns:
            List[Dict]: List of items from all pages
        """
        if not self.access_token:
            raise ValueError("No access token available. Please authenticate first.")
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        all_items = []
        next_link = f"https://graph.microsoft.com/v1.0{endpoint}"
        
        while next_link:
            try:
                # Handle both full URLs and relative endpoints
                if next_link.startswith("https://"):
                    url = next_link
                else:
                    url = f"https://graph.microsoft.com/v1.0{next_link}"
                
                response = requests.get(url, headers=headers, params=params)
                response.raise_for_status()
                
                data = response.json()
                
                # Extract items from current page
                if "value" in data:
                    all_items.extend(data["value"])
                
                # Check for next page
                next_link = data.get("@odata.nextLink", None)
                if next_link and next_link.startswith("https://graph.microsoft.com/v1.0"):
                    next_link = next_link.replace("https://graph.microsoft.com/v1.0", "")
                
                print(f"📄 Retrieved {len(data.get('value', []))} items from current page")
                
            except requests.exceptions.RequestException as e:
                print(f"❌ API request failed: {str(e)}")
                break
            except Exception as e:
                print(f"❌ Unexpected error: {str(e)}")
                break
        
        return all_items
    
    def fetch_inbox_emails(self) -> List[Dict[str, Any]]:
        """
        Fetch inbox emails containing 'wisp' in the subject.
        
        Returns:
            List[Dict]: List of inbox emails
        """
        print("📧 Fetching inbox emails containing 'wisp'...")
        
        params = {
            "$filter": "contains(subject, 'wisp')",
            "$select": "id,subject,from,toRecipients,receivedDateTime,bodyPreview,importance,isRead",
            "$orderby": "receivedDateTime desc"
        }
        
        # Use organization-wide endpoint with specific user
        user_email = "thyer@advantagewisp.com"
        endpoint = f"/users/{user_email}/messages"
        emails = self._make_graph_request(endpoint, params)
        print(f"✅ Retrieved {len(emails)} inbox emails")
        return emails
    
    def fetch_sent_emails(self) -> List[Dict[str, Any]]:
        """
        Fetch emails from the Sent Items folder.
        
        Returns:
            List[Dict]: List of sent emails
        """
        print("📤 Fetching sent emails...")
        
        params = {
            "$select": "id,subject,from,toRecipients,sentDateTime,bodyPreview,importance",
            "$orderby": "sentDateTime desc"
        }
        
        # Use organization-wide endpoint with specific user
        user_email = "thyer@advantagewisp.com"
        endpoint = f"/users/{user_email}/mailFolders('sentitems')/messages"
        emails = self._make_graph_request(endpoint, params)
        print(f"✅ Retrieved {len(emails)} sent emails")
        return emails
    
    def fetch_calendar_events(self) -> List[Dict[str, Any]]:
        """
        Fetch calendar events.
        
        Returns:
            List[Dict]: List of calendar events
        """
        print("📅 Fetching calendar events...")
        
        params = {
            "$select": "id,subject,start,end,location,bodyPreview,importance,isAllDay,recurrence",
            "$orderby": "start/dateTime desc"
        }
        
        # Use organization-wide endpoint with specific user
        user_email = "thyer@advantagewisp.com"
        endpoint = f"/users/{user_email}/events"
        events = self._make_graph_request(endpoint, params)
        print(f"✅ Retrieved {len(events)} calendar events")
        return events
    
    def fetch_all_data(self) -> Dict[str, Any]:
        """
        Fetch all data from Outlook and return as structured JSON.
        
        Returns:
            Dict: Combined data from all sources
        """
        if not self.authenticate():
            raise Exception("Authentication failed")
        
        print("\n🚀 Starting data retrieval...")
        
        # Fetch all data
        inbox_emails = self.fetch_inbox_emails()
        sent_emails = self.fetch_sent_emails()
        calendar_events = self.fetch_calendar_events()
        
        # Combine all data
        combined_data = {
            "retrieval_timestamp": datetime.now().isoformat(),
            "total_items": {
                "inbox_emails": len(inbox_emails),
                "sent_emails": len(sent_emails),
                "calendar_events": len(calendar_events)
            },
            "inbox_emails": inbox_emails,
            "sent_emails": sent_emails,
            "calendar_events": calendar_events
        }
        
        print(f"\n✅ Data retrieval complete!")
        print(f"   📧 Inbox emails: {len(inbox_emails)}")
        print(f"   📤 Sent emails: {len(sent_emails)}")
        print(f"   📅 Calendar events: {len(calendar_events)}")
        
        return combined_data
    
    def save_data_to_file(self, data: Dict[str, Any], filename: str = "outlook_data.json") -> None:
        """
        Save the retrieved data to a JSON file.
        
        Args:
            data (Dict): Data to save
            filename (str): Output filename
        """
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False, default=str)
            
            print(f"💾 Data saved to {filename}")
            
        except Exception as e:
            print(f"❌ Error saving data: {str(e)}")

def main():
    """
    Main function to demonstrate usage of the OutlookAuthenticator class.
    """
    print("🔐 Microsoft Outlook AI Agent - Authentication & Data Retrieval")
    print("=" * 60)
    
    # Import configuration
    try:
        from config import TENANT_ID, CLIENT_ID, CLIENT_SECRET
    except ImportError:
        print("❌ Error: config.py not found. Please copy config_template.py to config.py and fill in your credentials.")
        return
    except Exception as e:
        print(f"❌ Error importing configuration: {str(e)}")
        return
    
    try:
        # Create authenticator instance
        authenticator = OutlookAuthenticator(
            tenant_id=TENANT_ID,
            client_id=CLIENT_ID,
            client_secret=CLIENT_SECRET
        )
        
        # Fetch all data
        data = authenticator.fetch_all_data()
        
        # Save to file
        authenticator.save_data_to_file(data)
        
        print("\n🎉 All operations completed successfully!")
        
    except Exception as e:
        print(f"\n❌ Error in main execution: {str(e)}")
        print("\nPlease check your configuration and try again.")

if __name__ == "__main__":
    main()
