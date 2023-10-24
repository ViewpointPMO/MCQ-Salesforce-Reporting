# GAS Script for Salesforce Integration

This Google App Script (GAS) allows you to interact with Salesforce using OAuth2.0. The GAS script fetches data from Salesforce based on criteria from Google Sheets and updates another Google Sheet accordingly.

## Setup and Configuration

1. **Create a Salesforce Developer Org**  
   - Go to the Salesforce Developer website at [https://developer.salesforce.com/signup](https://developer.salesforce.com/signup).
   - Sign up and wait for an email from Salesforce. This email contains an activation link.
   - Click on the link to activate your account.

2. **Setup Connected App in Salesforce**
   - After logging into your developer environment, navigate to "Setup".
   - Create a new "Connected App" via **Apps** -> **App Manager**.
   - The connected app needs the following settings:
     - **OAuth Scopes**: 
       - Access and manage your data (api)
       - Perform requests at any time (refresh_token, offline_access)
     - **Callback URL**: It should point to the GAS's URL ending with `/usercallback`. For example, `https://script.google.com/macros/d/UNIQUE_ID/usercallback`.
     - **Settings**:
       - Require Secret for Web Server Flow: TRUE
       - Require Secret for Refresh Token Flow: TRUE

3. **Salesforce Authorization in GAS**
   - Inside the GAS editor:
     - Add `Oauth2` library and `Sheets` service
     - Enter the Consumer Key and Secret from the Salesforce Connected App and set up the necessary script and user properties (`SFClientID` and `SFClientSecret`, respectively).
     - Run the function `getSfService`.
     - Then run the function `showSidebar`.
   - **Note**: Before logging in, ensure you've cleared your cookies. This ensures you log into the actual production Salesforce environment and not the developer environment using your real Salesforce account.

## GAS Code Overview

The provided GAS script includes functions to:
- Interact with Salesforce's API.
- Update specific or all sheets in a Google Spreadsheet.
- Convert Salesforce's response to a format suitable for Google Sheets.
- Handle OAuth2.0 authentication and token refresh.

Key functions:
- `updateSingleSheetPrompted()`: Updates a single sheet based on user's input.
- `updateAllSheets()`: Updates all sheets defined in settings.
- `refreshToken()`: Refreshes Salesforce OAuth token when needed.

## Getting Started

1. Copy the GAS code into your Google App Script editor.
3. Follow the Salesforce setup and configuration steps mentioned above.
4. Use the custom Salesforce menu in Google Sheets before running the main script function for the first time.
