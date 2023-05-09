require('isomorphic-fetch');

const azure = require('@azure/identity');
const graph = require('@microsoft/microsoft-graph-client');
const authProviders =
    require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

let _settings = undefined;
let _clientSecretCredential = undefined;
let _appClient = undefined;

function ensureGraphForAppOnlyAuth(settings) {
    _settings = settings
    // Ensure settings isn't null
    if (!_settings)
        throw new Error('Settings cannot be undefined');

    if (!_clientSecretCredential) {
        _clientSecretCredential = new azure.ClientSecretCredential(
            _settings.tenantId,
            _settings.clientId,
            _settings.clientSecret
        );
    }

    if (!_appClient) {
        const authProvider = new authProviders.TokenCredentialAuthenticationProvider(
            _clientSecretCredential, {
            scopes: ['https://graph.microsoft.com/.default']
        });

        _appClient = graph.Client.initWithMiddleware({
            authProvider: authProvider
        });
    }
}

async function retrieveJoinedTeamsAsync(id) {
    return _appClient?.api('/users/' + id + '/joinedTeams')
        .get();
}
//Functions below are for scopes using application permissions 
async function retrieveUsersAsync() {
    return _appClient?.api('/users')
        .select(['displayName', 'id', 'mail'])
        .top(25)
        .orderby('displayName')
        .get();
}
async function retrieveUserEmail(id) {
    return _appClient?.api('/users/' + id)
        .select(['userPrincipalName'])
        .get();
}
module.exports = {
    ensureGraphForAppOnlyAuth,
    retrieveJoinedTeamsAsync,
    retrieveUserEmail,
    retrieveUsersAsync
}