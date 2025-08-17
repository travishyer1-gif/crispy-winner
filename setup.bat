@echo off
echo Microsoft Outlook AI Agent - Setup
echo =================================
echo.
echo Installing required Python packages...
pip install -r requirements.txt
echo.
echo Setup complete! You can now:
echo 1. Edit config.py with your credentials
echo 2. Run run_tests.bat to test your setup
echo 3. Run run_agent.bat to retrieve Outlook data
echo.
echo Press any key to exit...
pause >nul
