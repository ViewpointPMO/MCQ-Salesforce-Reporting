function updateTestSheet() {
  refreshToken();

  const soqlQuery = `SELECT FIELDS(ALL) FROM Issues__c LIMIT 10`;

  const sfService = getSfService();
  const oauthData = JSON.parse(PropertiesService.getUserProperties().getProperty('oauth2.salesforce'));
  let queryUrl = `${oauthData.instance_url}/services/data/v51.0/query?q=${encodeURIComponent(soqlQuery)}`;

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

  const ss = SpreadsheetApp.getActive();
  let responseSheet = ss.getSheetByName("TEST");
  if (!responseSheet) {
    // If the "TEST" sheet doesn't exist, create it
    responseSheet = ss.insertSheet("TEST");
  }

  // Clear existing data and write new data
  responseSheet.clear();
  responseSheet.getRange(1, 1, resultArray.length, resultArray[0].length).setValues(resultArray);
}
