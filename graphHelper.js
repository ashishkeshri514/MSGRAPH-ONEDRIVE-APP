// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <UserAuthConfigSnippet>
require('isomorphic-fetch');
const azure = require('@azure/identity');
const graph = require('@microsoft/microsoft-graph-client');
const authProviders =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

let _settings = undefined;
let _deviceCodeCredential = undefined;
let _userClient = undefined;

function initializeGraphForUserAuth(settings, deviceCodePrompt) {
  // Ensure settings isn't null
  if (!settings) {
    throw new Error('Settings cannot be undefined');
  }

  _settings = settings;

  _deviceCodeCredential = new azure.DeviceCodeCredential({
    clientId: settings.clientId,
    tenantId: settings.tenantId,
    userPromptCallback: deviceCodePrompt
  });

  const authProvider = new authProviders.TokenCredentialAuthenticationProvider(
    _deviceCodeCredential, {
      scopes: settings.graphUserScopes
    });

  _userClient = graph.Client.initWithMiddleware({
    authProvider: authProvider
  });
}
module.exports.initializeGraphForUserAuth = initializeGraphForUserAuth;



async function listFilesAsync() {
  // INSERT YOUR CODE HERE
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  return _userClient.api('/me/drive/root/children')
    .get();
}
module.exports.listFilesAsync = listFilesAsync;

async function downloadFileAsync(fileId) {
  try {
    if (!_userClient) {
      throw new Error('Graph has not been initialized for user auth');
    }

    const metadata = await _userClient
      .api(`/me/drive/items/${fileId}`)
      .get();
    
    // Fetch the file content using Graph API
    const contentResponse = await _userClient
      .api(`/me/drive/items/${fileId}/content`)
      .responseType('arraybuffer') // Ensure response type is set to 'arraybuffer'
      .get();

    return { metadata, content: contentResponse }; // Return both metadata and content
  } catch (error) {
    console.error('Error downloading file:', error.message);
    throw error; // Propagate the error to handle it in the caller function
  }
}
module.exports.downloadFileAsync = downloadFileAsync;

async function getUsersWithAccessAsync(fileId) {
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }
  const fileMetadata = await _userClient
      .api(`/me/drive/items/${fileId}`)
      .get();

  const filePermissions = await _userClient.api(`/me/drive/items/${fileId}/permissions`).get();

  return {filePermissions , fileMetadata};
}
module.exports.getUsersWithAccessAsync = getUsersWithAccessAsync;

