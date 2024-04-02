/**
 * Configures the service.
 */
function getService_() {
  const scriptProp = PropertiesService.getScriptProperties();
  const clientID = scriptProp.getProperty("OAUTH2_CLIENT_ID");
  const clientSecret = scriptProp.getProperty("OAUTH2_CLIENT_SECRET");

  if (!clientID || !clientSecret) {
    throw new Error(
      "Script properties must be set to run the script. Please set the 'OAUTH2_CLIENT_ID' and 'OAUTH2_CLIENT_SECRET' properties."
    );
  }

  return (
    OAuth2.createService("Mendeley")
      // Set the endpoint URLs.
      .setAuthorizationBaseUrl("https://api.mendeley.com/oauth/authorize")
      .setTokenUrl("https://api.mendeley.com/oauth/token")

      // Set the client ID and secret.
      .setClientId(clientID)
      .setClientSecret(clientSecret)

      // Set the name of the callback function that should be invoked to
      // complete the OAuth flow.
      .setCallbackFunction("authCallback")

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())

      // Set the scope. The "signature" scope is used for all endpoints in the
      // eSignature REST API.
      .setScope("all")
  );
}

/**
 * Handles the OAuth callback.
 */
function authCallback(request: object) {
  var service = getService_();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput("Success!");
  } else {
    return HtmlService.createHtmlOutput("Denied.");
  }
}

/**
 * Logs the redict URI to register in the Dropbox application settings.
 */
function logRedirectUri() {
  var service = getService_();
  Logger.log(service.getRedirectUri());
}
