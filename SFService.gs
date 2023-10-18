function getSfService() {
  const props = PropertiesService.getScriptProperties();
  return OAuth2.createService('salesforce')
    .setAuthorizationBaseUrl('https://login.salesforce.com/services/oauth2/authorize')
    .setTokenUrl('https://login.salesforce.com/services/oauth2/token')
    .setClientId(props.getProperty("SFClientID"))
    .setClientSecret(props.getProperty("SFClientSecret"))
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('api refresh_token');
}

function showSidebar() {
  const sfService = getSfService();
  if (!sfService.hasAccess()) {
    const authorizationUrl = sfService.getAuthorizationUrl();
    const template = HtmlService.createTemplate(
      `<a href="${authorizationUrl}" target="_blank">Authorize</a>. Reopen the sidebar when the authorization is complete.`
    );
    const page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  } else {
    SpreadsheetApp.getActive().toast('Authorization already done.');
  }
}

function authCallback(request) {
  const sfService = getSfService();
  const isAuthorized = sfService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}