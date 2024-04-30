function refreshToken(): string {
  var userProperties = PropertiesService.getUserProperties();
  var refresh_token = userProperties.getProperty("REFRESH_TOKEN");

  var scriptProperties = PropertiesService.getScriptProperties();
  var client_id = scriptProperties.getProperty("OAUTH2_CLIENT_ID");
  var client_secret = scriptProperties.getProperty("OAUTH2_CLIENT_SECRET");

  // Generate initial access token
  var response = UrlFetchApp.fetch("https://api.mendeley.com/oauth/token", {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    headers: {
      Authorization:
        "Basic " + Utilities.base64Encode(`${client_id}:${client_secret}`),
    },
    payload: {
      grant_type: "refresh_token",
      refresh_token: refresh_token,
      redirect_uri: "https://bibs4mendeley.pages.dev",
    },
  });

  var json = JSON.parse(response.getContentText());
  var current_time = Math.floor(Date.now() / 1000);
  userProperties.setProperty("ACCESS_TOKEN", json.access_token);
  userProperties.setProperty("REFRESH_TOKEN", json.refresh_token);
  userProperties.setProperty("EXPIRE_TIME", current_time + json.expires_in);
  return json.access_token;
}

function initAccessToken(code: string) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("OAUTH2_CODE", code);

  var scriptProperties = PropertiesService.getScriptProperties();
  var client_id = scriptProperties.getProperty("OAUTH2_CLIENT_ID");
  var client_secret = scriptProperties.getProperty("OAUTH2_CLIENT_SECRET");

  // Generate initial access token
  var response = UrlFetchApp.fetch("https://api.mendeley.com/oauth/token", {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    headers: {
      Authorization:
        "Basic " + Utilities.base64Encode(`${client_id}:${client_secret}`),
    },
    payload: {
      grant_type: "authorization_code",
      code: code,
      redirect_uri: "https://bibs4mendeley.pages.dev",
    },
  });

  var json = JSON.parse(response.getContentText());
  var current_time = Math.floor(Date.now() / 1000);
  var expires_time = current_time + json.expires_in - 30;
  userProperties.setProperty("ACCESS_TOKEN", json.access_token);
  userProperties.setProperty("REFRESH_TOKEN", json.refresh_token);
  userProperties.setProperty("EXPIRE_TIME", expires_time.toString());

  var ui = DocumentApp.getUi();
  ui.alert("Access token generated successfully!");
}

function getAccessToken(): string {
  var userProperties = PropertiesService.getUserProperties();
  var current_time = Math.floor(Date.now() / 1000);
  var expire_time = userProperties.getProperty("EXPIRE_TIME");
  var access_token = userProperties.getProperty("ACCESS_TOKEN");

  if (!expire_time || !access_token) {
    mendeleyLogin();
    throw new Error(
      "Access token not found! Please connect your Mendeley account."
    );
  }

  if (current_time > parseInt(expire_time)) {
    access_token = refreshToken();
  }
  return access_token;
}

function resetAccessToken() {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty("ACCESS_TOKEN");
  userProperties.deleteProperty("REFRESH_TOKEN");
  userProperties.deleteProperty("EXPIRE_TIME");
}
