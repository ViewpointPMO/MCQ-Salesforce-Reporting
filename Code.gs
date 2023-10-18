const FIELDS = 'Id, CaseNumber, Type, Support_Product_Portal__c, Module_Portal__c, Component_Portal__c, Status, Priority, Reason, Subject, Description, IsClosed, ClosedDate, IsEscalated, CreatedDate, Time_to_Initial_Response_Elapsed__c, Resolution_Time_Elapsed_days__c';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Salesforce')
    .addItem('Run Query', 'updateCases')
    .addToUi();
}

function updateCases() {
  refreshToken();
  makeRequestSoql();
  setReportDateRange();
}

function refreshToken() {
  const scriptProps = PropertiesService.getScriptProperties();
  const userProps = PropertiesService.getUserProperties();

  const clientId = scriptProps.getProperty("SFClientID");
  const clientSecret = scriptProps.getProperty("SFClientSecret");
  
  // Check if clientId or clientSecret are missing
  if (!clientId || !clientSecret) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.prompt('Configuration Missing', 'Please add SFClientID and SFClientSecret in script properties.', ui.ButtonSet.OK);
      if (response.getSelectedButton() == ui.Button.OK) {
          Logger.log(`User didn't provide the necessary info.`)
      }
      return;
  }

  let currentOAuthProps = userProps.getProperty('oauth2.salesforce');
  
  // If currentOAuthProps is missing or empty, initiate authentication flow
  if (!currentOAuthProps || Object.keys(JSON.parse(currentOAuthProps)).length === 0) {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Authentication Required', 'You need to authenticate with Salesforce before proceeding.', ui.ButtonSet.OK);
      getSfService();
      showSidebar();
      return;
  }

  currentOAuthProps = JSON.parse(currentOAuthProps);

  const refreshTokenEndpoint = `login.salesforce.com/services/oauth2/token?grant_type=refresh_token&client_id=${clientId}&client_secret=${clientSecret}&refresh_token=${currentOAuthProps.refresh_token}`;
  const response = UrlFetchApp.fetch(refreshTokenEndpoint, { method: 'POST' });

  const newOAuthProps = JSON.parse(response);
  Object.assign(currentOAuthProps, {
    signature: newOAuthProps.signature,
    access_token: newOAuthProps.access_token,
    issued_at: newOAuthProps.issued_at
  });

  userProps.setProperty('oauth2.salesforce', JSON.stringify(currentOAuthProps));
}

function getPastDate() {
  let today = new Date();
  let year = today.getFullYear() - 2;
  let month = today.getMonth();  // Months are 0-indexed in JavaScript.

  if (month === 0) {
    year -= 1;
    month = 11;  // December of the previous year.
  } else {
    month -= 1;  // Subtract 1 month.
  }

  // Convert single-digit months to two-digit format.
  let formattedMonth = (month + 1).toString().padStart(2, '0');

  return `${year}-${formattedMonth}-01T00:00:00Z`;
}

function makeRequestSoql() {
  const ss = SpreadsheetApp.getActive();
  const settingsSheet = ss.getSheetByName('Settings');
  const dataRange = settingsSheet.getDataRange().getValues();
  const sheetsUpdated = new Set(); // Track which sheets have been updated
  const numRows = settingsSheet.getLastRow();

  for (let i = 3; i < numRows; i++) {
    const rowData = dataRange[i];
    const parentSheetName = rowData[0];
    const supportProduct = rowData[1];
    const types = rowData[2];
    const modulePortal = rowData[3];
    const shouldUpdate = rowData[4];

    if (shouldUpdate !== "Yes") {
      continue;  // Skip the current iteration and move to the next row
    }

    let whereClause = `WHERE Status != 'Cancelled' AND CreatedDate > ${getPastDate()}`;
    if (supportProduct !== undefined) {
      whereClause += ` AND Support_Product_Portal__c = '${supportProduct}'`;
    }
    if (types) {
      const typeList = types.split(",").map(type => type.trim());
      const typeConditions = typeList.map(type => `Type = '${type}'`).join(" OR ");
      whereClause += ` AND (${typeConditions})`;
    }
    if (modulePortal) {
      whereClause += ` AND Module_Portal__c = '${modulePortal}'`;
    }

    const soqlTemplate = `SELECT ${FIELDS} FROM Case ${whereClause}`;
    const sfService = getSfService();
    const oauthData = JSON.parse(PropertiesService.getUserProperties().getProperty('oauth2.salesforce'));
    let queryUrl = `${oauthData.instance_url}/services/data/v51.0/query?q=${encodeURIComponent(soqlTemplate)}`;

    do {
      const response = UrlFetchApp.fetch(queryUrl, {
        method: "GET",
        headers: { "Authorization": `OAuth ${sfService.getAccessToken()}` },
        muteHttpExceptions: true
      });

      if (response.getResponseCode() !== 200) {
        Logger.log(response.getContentText());
        return;
      }

      const responseData = JSON.parse(response.getContentText());
      const resultArray = toArray(responseData.records);

      let responseSheet = ss.getSheetByName(parentSheetName);
      if (!responseSheet) {
        Logger.log(`Sheet ${parentSheetName} not found. Ensure the sheet exists.`);
        return;
      }

      if (!sheetsUpdated.has(parentSheetName)) {
        if (responseSheet.getLastRow() > 1) {
          responseSheet.getRange(2, 1, responseSheet.getLastRow() - 1, responseSheet.getLastColumn()).clearContent();
        }
        sheetsUpdated.add(parentSheetName);
      }

      if (resultArray.length > 1) {
        const lastRow = responseSheet.getLastRow();
        responseSheet.getRange(lastRow + 1, 1, resultArray.length - 1, resultArray[0].length).setValues(resultArray.slice(1));
        settingsSheet.getRange(i + 1, 7).setValue(responseData.totalSize + 0);
      } else {
        Logger.log(`No results to write for sheet ${parentSheetName} - (Product: ${supportProduct})`);
      }

      queryUrl = responseData.nextRecordsUrl ? `${oauthData.instance_url}${responseData.nextRecordsUrl}` : null;

    } while (queryUrl);
  }
}

function toArray(items) {
  // Check if items is empty or the first item is undefined/null
  if (!items.length || !items[0]) {
    return [];
  }

  const headers = Object.keys(items[0]).filter(key => key !== "attributes").concat(["FormattedDate"]);
  const data = items.map(item => {
    let values = headers.slice(0, -1).map(header => item[header]);
    let formattedDate = formatCreatedDate(item.CreatedDate); // Add formatted date at the end.
    values.push(formattedDate);
    return values;
  });

  return [headers, ...data];
}

function formatCreatedDate(dateStr) {
  let date = new Date(dateStr);
  let month = (date.getMonth() + 1).toString().padStart(2, '0'); // Months are 0-indexed.
  let year = date.getFullYear();
  return `${month}-${year}`;
}


function getFormattedCurrentDate() {
    let today = new Date();
    let day = today.getDate().toString().padStart(2, '0');
    let month = (today.getMonth() + 1).toString().padStart(2, '0'); // Months are 0-indexed.
    let year = today.getFullYear();
    return `${month}-${day}-${year}`;
}

function setReportDateRange() {
    const ss = SpreadsheetApp.getActive();
    const settingsSheet = ss.getSheetByName('Settings');
    const summarySheet = ss.getSheetByName('Summary'); // Get 'Summary' sheet
    
    let currentDate = new Date();
    let startDate = new Date(currentDate.getFullYear() - 2, currentDate.getMonth() - 1, 1); // 2 years and 1 month ago from the current date.
    let endDate = new Date(); // Current date.
    
    // Reset the hours, minutes, seconds, and milliseconds
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(0, 0, 0, 0);

    // Set the two dates in the 'Settings' sheet at D1 and E1
    settingsSheet.getRange('E1').setValue(startDate);
    settingsSheet.getRange('F1').setValue(endDate);

    // Set the two dates in the 'Summary' sheet at B1 and C1
    summarySheet.getRange('B1').setValue(startDate);
    summarySheet.getRange('C1').setValue(endDate);

    // Generate and write the list of month start dates between startDate and endDate
    let monthDates = getMonthStartDates(startDate, endDate);
    summarySheet.getRange(5, 1, monthDates.length, 1).setValues(monthDates.map(date => [date]));
}

function getMonthStartDates(startDate, endDate) {
    let dates = [];
    let currentDate = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
    
    while (currentDate <= endDate) {
        let month = (currentDate.getMonth() + 1).toString().padStart(2, '0'); // Months are 0-indexed.
        let year = currentDate.getFullYear();
        dates.push(`${month}-${year}`);
        
        currentDate.setMonth(currentDate.getMonth() + 1); // Move to the next month.
    }
    
    return dates;
}