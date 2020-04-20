import {UserAgentApplication} from 'msal'

const msalConfig = {
    auth: {
        clientId: "<YOUR_CLIENT_ID>", //client id from app registration `application (client) id`
        authority: "https://login.microsoftonline.com/common", //default authority is `https://login.microsoftonline.com/common`
        redirectUri: "<YOUR_REDIRECT_URI>", //must be same with App registration Redirect URI, otherwise error
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    }
};

export const loginRequest = {
    scopes: ["openid", "profile", "User.Read"]
};

export const msal = new UserAgentApplication(msalConfig);

export const graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me"
};

export default {loginRequest, msal, graphConfig}

