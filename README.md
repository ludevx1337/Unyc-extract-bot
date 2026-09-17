# UNYC Atlas Automation Bot

This Python script automates the login process to UNYC Atlas, navigates to client pages, and extracts RIO codes from client phone numbers using Playwright.

## Features

- Automated login to https://atlas.unyc.io/
- Handles 2FA authentication (manual input required)
- Navigates to client search page
- Reads client names from Excel file (Feuil1 sheet)
- Performs client searches automatically
- **NEW**: Extracts RIO codes from client phone numbers
- **NEW**: Saves RIO codes back to Excel file with phone number mapping
- Processes multiple phone numbers per client
- Keeps browser open for manual interaction

## Setup

1. **Install dependencies:**
   ```bash
   python setup.py
   ```
   
   Or manually:
   ```bash
   pip install -r requirements.txt
   playwright install chromium
   ```

## Usage

1. **Run the automation script:**
   ```bash
   python unyc_automation.py
   ```

3. **The script will:**
   - Open a browser window
   - Navigate to the UNYC Atlas login page
   - Fill in the credentials automatically
   - Submit the login form
   - If 2FA is required, pause and wait for manual input
   - Navigate to the client search page
   - Read client names from Excel file (Feuil1 sheet)
   - For each client:
     - Search and navigate to client page
     - Extract RIO codes from all phone numbers
     - Save RIO codes to Excel file
   - Provide detailed summary report

## Excel File Requirements

**File Name**: `lmunyc.xlsx`  
**Sheet Name**: `Feuil1`  
**Required Column**: `Client`  
**Auto-Created Column**: `RIO` (will be created automatically)

### Excel Structure:

| Client | RIO |
|--------|-----|
| ABC DELIGHT DENTAIRE | 0607371063:RIO123456; 0698741258:RIO789012 |
| ANOTHER CLIENT | 0612345678:RIO345678 |

**RIO Format**: `PhoneNumber:RIOCode; PhoneNumber2:RIOCode2`

## RIO Extraction Process

For each client, the script will:
1. Navigate to the client's page
2. Access the "Telephonie" section
3. Set license display to "Tous" (all)
4. Find all phone numbers in the table
5. For each phone number:
   - Expand the details row
   - Click on "Lignes mobiles"
   - Click on the specific phone number
   - Extract the RIO code from the form
   - Save to Excel file
6. Return to client search and continue with next client

## 2FA Handling

If the system requires 2FA:
1. The script will pause and display a message
2. Manually enter the 2FA code in the browser
3. Press Enter in the console to continue

## Browser Behavior

- The browser runs in visible mode (not headless) for easier debugging
- Operations are slowed down for visibility
- The browser remains open after automation for manual interaction
- Press Enter in the console to close the browser

## Error Handling

The script includes error handling for:
- Network timeouts
- Missing elements
- Login failures
- 2FA detection

## Files

- `unyc_automation.py` - Main automation script
- `setup.py` - Setup and installation script
- `requirements.txt` - Python dependencies
- `README.md` - This documentation file