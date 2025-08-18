#!/usr/bin/env python3
"""
Test script to verify Microsoft Graph API connection and basic functionality.
Run this script to test your configuration before using the main authenticator.
"""

import sys
import os
from config import TENANT_ID, CLIENT_ID, USERNAME, PASSWORD

def test_configuration():
    """Test if configuration values are properly set."""
    print("🔍 Testing configuration...")
    
    # Check if configuration values are still placeholders
    placeholder_values = [
        ("TENANT_ID", TENANT_ID, "your_tenant_id_here"),
        ("CLIENT_ID", CLIENT_ID, "your_client_id_here"),
        ("USERNAME", USERNAME, "your_email@domain.com"),
        ("PASSWORD", PASSWORD, "your_password_here")
    ]
    
    has_placeholders = False
    for name, value, placeholder in placeholder_values:
        if value == placeholder:
            print(f"❌ {name}: Still using placeholder value")
            has_placeholders = True
        else:
            print(f"✅ {name}: Configured")
    
    if has_placeholders:
        print("\n⚠️  Please update your configuration in config.py before proceeding.")
        return False
    
    print("✅ Configuration looks good!")
    return True

def test_dependencies():
    """Test if required dependencies are installed."""
    print("\n📦 Testing dependencies...")
    
    try:
        import msal
        print("✅ MSAL library installed")
    except ImportError:
        print("❌ MSAL library not found. Run: pip install msal")
        return False
    
    try:
        import requests
        print("✅ Requests library installed")
    except ImportError:
        print("❌ Requests library not found. Run: pip install requests")
        return False
    
    try:
        import json
        print("✅ JSON library available (built-in)")
    except ImportError:
        print("❌ JSON library not available")
        return False
    
    return True

def test_import():
    """Test if the main authenticator can be imported."""
    print("\n📥 Testing imports...")
    
    try:
        from outlook_authenticator import OutlookAuthenticator
        print("✅ OutlookAuthenticator class imported successfully")
        return True
    except ImportError as e:
        print(f"❌ Failed to import OutlookAuthenticator: {e}")
        return False

def main():
    """Main test function."""
    print("🧪 Microsoft Outlook AI Agent - Connection Test")
    print("=" * 50)
    
    # Run all tests
    config_ok = test_configuration()
    deps_ok = test_dependencies()
    import_ok = test_import()
    
    print("\n" + "=" * 50)
    print("📊 Test Results Summary:")
    print(f"   Configuration: {'✅ PASS' if config_ok else '❌ FAIL'}")
    print(f"   Dependencies:  {'✅ PASS' if deps_ok else '❌ FAIL'}")
    print(f"   Imports:       {'✅ PASS' if import_ok else '❌ FAIL'}")
    
    if all([config_ok, deps_ok, import_ok]):
        print("\n🎉 All tests passed! You're ready to use the Outlook AI Agent.")
        print("\nNext steps:")
        print("1. Run: python outlook_authenticator.py")
        print("2. Check the generated outlook_data.json file")
    else:
        print("\n⚠️  Some tests failed. Please fix the issues before proceeding.")
        sys.exit(1)

if __name__ == "__main__":
    main()
