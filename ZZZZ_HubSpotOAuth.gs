/**
 * HubSpot OAuth support for the Monthly Quality reporting package.
 *
 * This file intentionally sorts after HubSpot.gs so these authentication
 * functions override the static-token versions in that file.
 *
 * No OAuth secrets or tokens belong in GitHub. Store them in Apps Script
 * Project Settings > Script Properties.
 *
 * Required for OAuth setup:
 *   HUBSPOT_OAUTH_CLIENT_ID
 *   HUBSPOT_OAUTH_CLIENT_SECRET
 *
 * Created automatically after authorization:
 *   HUBSPOT_OAUTH_REFRESH_TOKEN
 *   HUBSPOT_OAUTH_ACCESS_TOKEN
 *   HUBSPOT_OAUTH_ACCESS_TOKEN_EXPIRES_AT
 *
 * Optional:
 *   HUBSPOT_OAUTH_REDIRECT_URI
 *
 * If HUBSPOT_OAUTH_REDIRECT_URI is omitted, the script uses the deployed
 * Apps Script web-app URL returned by ScriptApp.getService().getUrl().
 */

const HUBSPOT_OAUTH_CONFIG_ = Object.freeze({
  authorizeUrl: 'https://app.hubspot.com/oauth/authorize',
  tokenUrl: 'https://api.hubapi.com/oauth/2026-03/token',
  clientIdProperty: 'HUBSPOT_OAUTH_CLIENT_ID',
  clientSecretProperty: 'HUBSPOT_OAUTH_CLIENT_SECRET',
  redirectUriProperty: 'HUBSPOT_OAUTH_REDIRECT_URI',
  refreshTokenProperty: 'HUBSPOT_OAUTH_REFRESH_TOKEN',
  accessTokenProperty: 'HUBSPOT_OAUTH_ACCESS_TOKEN',
  accessTokenExpiresAtProperty: 'HUBSPOT_OAUTH_ACCESS_TOKEN_EXPIRES_AT',
  stateProperty: 'HUBSPOT_OAUTH_STATE',
  stateCreatedAtProperty: 'HUBSPOT_OAUTH_STATE_CREATED_AT',
  stateLifetimeMs: 15 * 60 * 1000,
  refreshBufferMs: 2 * 60 * 1000,
  scopes: Object.freeze([
    'crm.objects.tickets.read',
    'crm.objects.owners.read'
  ])
});

/**
 * Run after deploying this Apps Script project as a Web App.
 * Returns the HubSpot authorization URL and logs it.
 */
function getHubSpotOAuthAuthorizationUrl() {
  const properties = PropertiesService.getScriptProperties();
  const clientId = requireHubSpotOAuthProperty_(
    HUBSPOT_OAUTH_CONFIG_.clientIdProperty
  );
  requireHubSpotOAuthProperty_(HUBSPOT_OAUTH_CONFIG_.clientSecretProperty);

  const redirectUri = getHubSpotOAuthRedirectUri_();
  const state = Utilities.getUuid();
  properties.setProperty(HUBSPOT_OAUTH_CONFIG_.stateProperty, state);
  properties.setProperty(
    HUBSPOT_OAUTH_CONFIG_.stateCreatedAtProperty,
    String(Date.now())
  );

  const url = HUBSPOT_OAUTH_CONFIG_.authorizeUrl +
    '?client_id=' + encodeURIComponent(clientId) +
    '&redirect_uri=' + encodeURIComponent(redirectUri) +
    '&scope=' + encodeURIComponent(HUBSPOT_OAUTH_CONFIG_.scopes.join(' ')) +
    '&state=' + encodeURIComponent(state);

  Logger.log('HubSpot OAuth authorization URL: ' + url);
  return {
    authorizationUrl: url,
    redirectUri: redirectUri,
    scopes: HUBSPOT_OAUTH_CONFIG_.scopes.slice()
  };
}

/**
 * Web-app callback for HubSpot OAuth.
 *
 * Deploy the Apps Script project as a Web App, add that /exec URL as an
 * allowed redirect URL on the HubSpot app, then open the URL returned by
 * getHubSpotOAuthAuthorizationUrl(). HubSpot redirects here with ?code=...
 * and this function exchanges the code and stores the refresh token.
 */
function doGet(e) {
  try {
    const parameters = (e && e.parameter) || {};

    if (parameters.error) {
      throw new Error(
        'HubSpot authorization failed: ' + parameters.error +
        (parameters.error_description
          ? ' - ' + parameters.error_description
          : '')
      );
    }

    if (!parameters.code) {
      return HtmlService.createHtmlOutput(
        '<h2>HubSpot OAuth callback is ready.</h2>' +
        '<p>Run <code>getHubSpotOAuthAuthorizationUrl()</code> in Apps Script ' +
        'and open the URL it returns.</p>'
      );
    }

    validateHubSpotOAuthState_(parameters.state || '');
    const tokenResult = exchangeHubSpotOAuthAuthorizationCode_(
      parameters.code,
      getHubSpotOAuthRedirectUri_()
    );

    PropertiesService.getScriptProperties().deleteProperty(
      HUBSPOT_OAUTH_CONFIG_.stateProperty
    );
    PropertiesService.getScriptProperties().deleteProperty(
      HUBSPOT_OAUTH_CONFIG_.stateCreatedAtProperty
    );

    return HtmlService.createHtmlOutput(
      '<h2>HubSpot connected successfully.</h2>' +
      '<p>The refresh token is stored in Apps Script Script Properties. ' +
      'You can close this window and run the monthly quality report.</p>' +
      '<p>Access token expires in approximately ' +
      Number(tokenResult.expires_in || 0) + ' seconds; the script will refresh it automatically.</p>'
    );
  } catch (error) {
    Logger.log('HubSpot OAuth callback error: ' + error.stack);
    return HtmlService.createHtmlOutput(
      '<h2>HubSpot connection failed.</h2><pre>' +
      escapeHubSpotOAuthHtml_(String(error && error.message || error)) +
      '</pre>'
    );
  }
}

/**
 * Safe status check. Does not return client secrets, refresh tokens, or access
 * token values.
 */
function getHubSpotOAuthStatus() {
  const properties = PropertiesService.getScriptProperties();
  const expiresAt = Number(
    properties.getProperty(
      HUBSPOT_OAUTH_CONFIG_.accessTokenExpiresAtProperty
    ) || 0
  );

  const result = {
    clientIdConfigured: Boolean(
      properties.getProperty(HUBSPOT_OAUTH_CONFIG_.clientIdProperty)
    ),
    clientSecretConfigured: Boolean(
      properties.getProperty(HUBSPOT_OAUTH_CONFIG_.clientSecretProperty)
    ),
    refreshTokenConfigured: Boolean(
      properties.getProperty(HUBSPOT_OAUTH_CONFIG_.refreshTokenProperty)
    ),
    accessTokenCached: Boolean(
      properties.getProperty(HUBSPOT_OAUTH_CONFIG_.accessTokenProperty)
    ),
    accessTokenExpiresAt: expiresAt
      ? new Date(expiresAt).toISOString()
      : null,
    redirectUri: getHubSpotOAuthRedirectUriOrBlank_(),
    scopes: HUBSPOT_OAUTH_CONFIG_.scopes.slice()
  };

  Logger.log(JSON.stringify(result));
  return result;
}

/**
 * Forces a token refresh and then performs the existing connection test.
 */
function refreshAndValidateHubSpotOAuth() {
  refreshHubSpotOAuthAccessToken_(true);
  return validateHubSpotQualityConnection();
}

/**
 * Override from HubSpot.gs. Prefer OAuth when it has been configured; retain
 * HUBSPOT_ACCESS_TOKEN as a fallback for a future Service Key/private-app token.
 */
function getHubSpotQualityToken_() {
  const properties = PropertiesService.getScriptProperties();
  const clientId = String(
    properties.getProperty(HUBSPOT_OAUTH_CONFIG_.clientIdProperty) || ''
  ).trim();
  const clientSecret = String(
    properties.getProperty(HUBSPOT_OAUTH_CONFIG_.clientSecretProperty) || ''
  ).trim();
  const refreshToken = String(
    properties.getProperty(HUBSPOT_OAUTH_CONFIG_.refreshTokenProperty) || ''
  ).trim();

  const anyOAuthSetting = Boolean(clientId || clientSecret || refreshToken);
  if (anyOAuthSetting) {
    if (!clientId || !clientSecret) {
      throw new Error(
        'HubSpot OAuth setup is incomplete. Set both ' +
        HUBSPOT_OAUTH_CONFIG_.clientIdProperty + ' and ' +
        HUBSPOT_OAUTH_CONFIG_.clientSecretProperty + ' in Script Properties.'
      );
    }
    if (!refreshToken) {
      throw new Error(
        'HubSpot OAuth has not been authorized yet. Deploy the Apps Script ' +
        'project as a Web App, add its /exec URL to the HubSpot app redirect URLs, ' +
        'then run getHubSpotOAuthAuthorizationUrl() and open the returned URL.'
      );
    }
    return refreshHubSpotOAuthAccessToken_(false);
  }

  return String(
    properties.getProperty(HUBSPOT_QUALITY_CONFIG_.tokenProperty) || ''
  ).trim();
}

/** Override from HubSpot.gs with OAuth-aware error text. */
function requireHubSpotQualityToken_() {
  const token = getHubSpotQualityToken_();
  if (!token) {
    throw new Error(
      'HubSpot authentication is not configured. Configure OAuth using ' +
      HUBSPOT_OAUTH_CONFIG_.clientIdProperty + ' / ' +
      HUBSPOT_OAUTH_CONFIG_.clientSecretProperty +
      ', or store a valid Service Key/private-app access token in ' +
      HUBSPOT_QUALITY_CONFIG_.tokenProperty + '.'
    );
  }
  return token;
}

/**
 * Override from HubSpot.gs. On an API 401, force one OAuth refresh and retry.
 */
function hubSpotQualityRequestJson_(method, path, payload) {
  let refreshedAfter401 = false;

  for (let attempt = 1; attempt <= HUBSPOT_QUALITY_CONFIG_.maxAttempts; attempt++) {
    const token = requireHubSpotQualityToken_();
    const options = {
      method: method,
      headers: {
        Authorization: 'Bearer ' + token,
        Accept: 'application/json'
      },
      muteHttpExceptions: true
    };

    if (payload !== null && payload !== undefined) {
      options.contentType = 'application/json';
      options.payload = JSON.stringify(payload);
    }

    const response = UrlFetchApp.fetch(
      HUBSPOT_QUALITY_CONFIG_.baseUrl + path,
      options
    );
    const status = response.getResponseCode();
    const body = response.getContentText();

    if (status >= 200 && status < 300) {
      return body ? JSON.parse(body) : {};
    }

    if (status === 401 && hubSpotOAuthIsFullyConfigured_() && !refreshedAfter401) {
      refreshedAfter401 = true;
      refreshHubSpotOAuthAccessToken_(true);
      continue;
    }

    if (status === 429 && attempt < HUBSPOT_QUALITY_CONFIG_.maxAttempts) {
      const headers = response.getAllHeaders();
      const retryAfter = Number(
        headers['Retry-After'] || headers['retry-after'] || 2
      );
      Utilities.sleep(Math.max(1, retryAfter) * 1000);
      continue;
    }

    if (status >= 500 && status <= 599 &&
        attempt < HUBSPOT_QUALITY_CONFIG_.maxAttempts) {
      Utilities.sleep(attempt * 2000);
      continue;
    }

    throw new Error(
      'HubSpot API request failed (' + status + '): ' + body
    );
  }

  throw new Error('HubSpot API request failed after all retry attempts.');
}

function refreshHubSpotOAuthAccessToken_(forceRefresh) {
  const properties = PropertiesService.getScriptProperties();
  const existingAccessToken = String(
    properties.getProperty(HUBSPOT_OAUTH_CONFIG_.accessTokenProperty) || ''
  ).trim();
  const expiresAt = Number(
    properties.getProperty(
      HUBSPOT_OAUTH_CONFIG_.accessTokenExpiresAtProperty
    ) || 0
  );

  if (
    !forceRefresh &&
    existingAccessToken &&
    expiresAt > Date.now() + HUBSPOT_OAUTH_CONFIG_.refreshBufferMs
  ) {
    return existingAccessToken;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const tokenAfterLock = String(
      properties.getProperty(HUBSPOT_OAUTH_CONFIG_.accessTokenProperty) || ''
    ).trim();
    const expiryAfterLock = Number(
      properties.getProperty(
        HUBSPOT_OAUTH_CONFIG_.accessTokenExpiresAtProperty
      ) || 0
    );

    if (
      !forceRefresh &&
      tokenAfterLock &&
      expiryAfterLock > Date.now() + HUBSPOT_OAUTH_CONFIG_.refreshBufferMs
    ) {
      return tokenAfterLock;
    }

    const tokenResult = requestHubSpotOAuthToken_({
      grant_type: 'refresh_token',
      client_id: requireHubSpotOAuthProperty_(
        HUBSPOT_OAUTH_CONFIG_.clientIdProperty
      ),
      client_secret: requireHubSpotOAuthProperty_(
        HUBSPOT_OAUTH_CONFIG_.clientSecretProperty
      ),
      refresh_token: requireHubSpotOAuthProperty_(
        HUBSPOT_OAUTH_CONFIG_.refreshTokenProperty
      )
    });

    storeHubSpotOAuthTokenResult_(tokenResult);
    return String(tokenResult.access_token || '').trim();
  } finally {
    lock.releaseLock();
  }
}

function exchangeHubSpotOAuthAuthorizationCode_(authorizationCode, redirectUri) {
  const tokenResult = requestHubSpotOAuthToken_({
    grant_type: 'authorization_code',
    client_id: requireHubSpotOAuthProperty_(
      HUBSPOT_OAUTH_CONFIG_.clientIdProperty
    ),
    client_secret: requireHubSpotOAuthProperty_(
      HUBSPOT_OAUTH_CONFIG_.clientSecretProperty
    ),
    redirect_uri: redirectUri,
    code: authorizationCode
  });

  storeHubSpotOAuthTokenResult_(tokenResult);
  return tokenResult;
}

function requestHubSpotOAuthToken_(formData) {
  const response = UrlFetchApp.fetch(HUBSPOT_OAUTH_CONFIG_.tokenUrl, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: formData,
    muteHttpExceptions: true,
    headers: {
      Accept: 'application/json'
    }
  });

  const status = response.getResponseCode();
  const body = response.getContentText();
  let parsed = {};
  try {
    parsed = body ? JSON.parse(body) : {};
  } catch (error) {
    parsed = {};
  }

  if (status < 200 || status >= 300) {
    throw new Error(
      'HubSpot OAuth token request failed (' + status + '): ' + body
    );
  }

  if (!parsed.access_token) {
    throw new Error(
      'HubSpot OAuth token response did not include an access_token.'
    );
  }

  return parsed;
}

function storeHubSpotOAuthTokenResult_(tokenResult) {
  const properties = PropertiesService.getScriptProperties();
  const expiresInSeconds = Math.max(60, Number(tokenResult.expires_in || 1800));
  const expiresAt = Date.now() + expiresInSeconds * 1000;

  properties.setProperty(
    HUBSPOT_OAUTH_CONFIG_.accessTokenProperty,
    String(tokenResult.access_token)
  );
  properties.setProperty(
    HUBSPOT_OAUTH_CONFIG_.accessTokenExpiresAtProperty,
    String(expiresAt)
  );

  // HubSpot may rotate refresh tokens. Always persist the latest token returned.
  if (tokenResult.refresh_token) {
    properties.setProperty(
      HUBSPOT_OAUTH_CONFIG_.refreshTokenProperty,
      String(tokenResult.refresh_token)
    );
  }
}

function hubSpotOAuthIsFullyConfigured_() {
  const properties = PropertiesService.getScriptProperties();
  return Boolean(
    String(properties.getProperty(HUBSPOT_OAUTH_CONFIG_.clientIdProperty) || '').trim() &&
    String(properties.getProperty(HUBSPOT_OAUTH_CONFIG_.clientSecretProperty) || '').trim() &&
    String(properties.getProperty(HUBSPOT_OAUTH_CONFIG_.refreshTokenProperty) || '').trim()
  );
}

function validateHubSpotOAuthState_(returnedState) {
  const properties = PropertiesService.getScriptProperties();
  const expectedState = String(
    properties.getProperty(HUBSPOT_OAUTH_CONFIG_.stateProperty) || ''
  );
  const createdAt = Number(
    properties.getProperty(HUBSPOT_OAUTH_CONFIG_.stateCreatedAtProperty) || 0
  );

  if (!expectedState || !returnedState || expectedState !== returnedState) {
    throw new Error('HubSpot OAuth state validation failed. Start authorization again.');
  }
  if (!createdAt || Date.now() - createdAt > HUBSPOT_OAUTH_CONFIG_.stateLifetimeMs) {
    throw new Error('HubSpot OAuth authorization request expired. Start authorization again.');
  }
}

function getHubSpotOAuthRedirectUri_() {
  const redirectUri = getHubSpotOAuthRedirectUriOrBlank_();
  if (!redirectUri) {
    throw new Error(
      'No HubSpot OAuth redirect URI is available. Deploy this Apps Script ' +
      'project as a Web App, or set HUBSPOT_OAUTH_REDIRECT_URI in Script Properties.'
    );
  }
  return redirectUri;
}

function getHubSpotOAuthRedirectUriOrBlank_() {
  const configured = String(
    PropertiesService.getScriptProperties().getProperty(
      HUBSPOT_OAUTH_CONFIG_.redirectUriProperty
    ) || ''
  ).trim();
  if (configured) return configured;

  try {
    return String(ScriptApp.getService().getUrl() || '').trim();
  } catch (error) {
    return '';
  }
}

function requireHubSpotOAuthProperty_(propertyName) {
  const value = String(
    PropertiesService.getScriptProperties().getProperty(propertyName) || ''
  ).trim();
  if (!value) {
    throw new Error(
      'Missing Apps Script Script Property: ' + propertyName
    );
  }
  return value;
}

function escapeHubSpotOAuthHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
