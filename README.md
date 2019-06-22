# Call Yammer API from a JavaScript Single Page Application using msal.js
===========================================================

| [Getting Started](https://docs.microsoft.com/en-us/azure/active-directory/develop/guidedsetups/active-directory-javascriptspa)| [AAD Docs](https://aka.ms/aaddevv2) | [Library Reference](https://htmlpreview.github.io/?https://raw.githubusercontent.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-core/docs/classes/_useragentapplication_.useragentapplication.html) | [Support](README.md#community-help-and-support) | [Samples](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki/Samples)
| --- | --- | --- | --- | --- |


The MSAL library for JavaScript enables client-side JavaScript web applications, running in a web browser, to authenticate users using [Azure AD](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-overview) work and school accounts (AAD).

## Installation
Via NPM:

    npm install msal

Via CDN:

    <!-- Latest compiled and minified JavaScript -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/bluebird/3.3.4/bluebird.min.js"></script>
    <script src="https://secure.aadcdn.microsoftonline-p.com/lib/1.0.0/js/msal.js"></script>

See here for more details on [supported browsers and known compatability issues](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki/FAQs#q4-what-browsers-is-msaljs-supported-on).

<!-- ## Go [here](https://docs.microsoft.com/azure/active-directory/develop/guidedsetups/active-directory-javascriptspa) for information about this code sample and how to configure it -->

## OAuth 2.0 and the Implicit Flow
Msal implements the [Implicit Grant Flow](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-implicit-grant-flow), as defined by the OAuth 2.0 protocol and is [OpenID](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-protocols-oidc) compliant.

## Usage
The example below walks you through how to login a user and acquire a token to be used for Yammer APIs.

#### Prerequisite

Before using MSAL.js you will need to [register an application in Azure AD](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app) to get a valid `clientId` for configuration, and to register the routes that your app will accept redirect traffic on.

1. #### Register the AAD app on Azure portal
While registering the app, make sure these steps are followed: 
- Register your App by going to **Azure Active Directory** on the left side menu, followed by **App registrations**. Then, select **New registration**.
- Give your app a meaningful name.
- Select **Accounts in this organizational directory only** if you want only your network to have access to your app.
- Set your redirect URI.
- Next, you should see the App overview page with information such as **Application ID** and other things. 
- Go to **API permissions** and then **Add a permission**. Select **Yammer**. Mark the checkbox "user_impersonation" and then click **Add permission**. This step is extremely imporant in order to Authenticate and talk to the Yammer APIs. Granting admin consent will allow any user in the Tenant to hit the API.
- Go to **Authentication** and enable the **Access tokens** and **ID tokens** under the **Advanced Settings > Implicit Grant**. Hit Save.

2. #### Connect your AAD app to your SPA

In your JavaScript component, initialize the **configuration** variable with your AAD app details like this.

```JavaScript
var msalConfig = {
        auth: {
            clientId: <your client id>, //This is your client ID from AAD
            authority: "https://login.microsoftonline.com/<your tenant name>" //This is your tenant info

        },
        cache: {
            cacheLocation: "localStorage",
            storeAuthStateInCookie: true
        },
        resource: "https://api.yammer.com"
    };

    var yammerConfig = {
        yammerUrl: "https://api.yammer.com/api/v1/"
    };

    //This will allow the APP to call the Yammer APIs.
    var requestObj = {
        scopes: ["https://api.yammer.com/user_impersonation"]
    };

    var myMSALObj = new Msal.UserAgentApplication(msalConfig); //create a new MSAL object. 
```

3. #### Login the User

In order to call the the Yammer APIs, users need to login and acquire an AAD access token. Your app must login the user with either the `loginPopup` or the `loginRedirect` method to establish user context. 

When the login methods are called and the authentication of the user is completed by the Azure AD service, an [id token](https://docs.microsoft.com/en-us/azure/active-directory/develop/id-tokens) is returned which is used to identify the user with some basic information.

When you login a user, you can pass in scopes that the user can pre consent to on login. Please note that consenting to scopes on login, does not return an access_token for these scopes, but gives you the opportunity to obtain a token silently with these scopes passed in, with no further interaction from the user. Azure AD would grant an idToken based on the scope, which for us is **api.yammer.com/user_impersonation**. 

```JavaScript
function signIn() {
        myMSALObj.loginPopup(requestObj).then(function (loginResponse) {
            //Successful login
            showWelcomeMessage();
            //Call Yammer APIs using the token in the response
            acquireTokenPopupAndcallYammerAPI();
        }).catch(function (error) {
            //Please check the console for errors
            console.log(error);
        });
    }
```

4. #### Acquiring an access token and calling Yammer API

In MSAL, you can get access tokens for the APIs your app needs to call using the acquireTokenSilent method which makes a silent request(without prompting the user with UI) to Azure AD to obtain an access token. The Azure AD service then returns an access token containing the user consented scopes to allow your app to securely call the API.

You can use acquireTokenRedirect or acquireTokenPopup to initiate interactive requests, although, it is best practice to only show interactive experiences if you are unable to obtain a token silently due to interaction required errors. If you are using an interactive token call, it must match the login method used in your application. (loginPopup=> acquireTokenPopup, loginRedirect => acquireTokenRedirect).

If the acquireTokenSilent call fails with an error of type InteractionRequiredAuthError you will need to initiate an interactive request. This could happen for many reasons including scopes that have been revoked, expired tokens, or password changes.

acquireTokenSilent will look for a valid token in the cache, and if it is close to expiring or does not exist, will automatically try to refresh it for you.

This is an example of how we can acquire an Azure AD token and upon success call the Yammer API. Other API endpoints can be found on the [Yammer API page](https://developer.yammer.com/docs/rest-api-rate-limits). 

```JavaScript
function acquireTokenPopupAndcallYammerAPI() {
        //Always start with acquireTokenSilent to obtain a token in the signed in user from cache
        myMSALObj.acquireTokenSilent(requestObj).then(function (tokenResponse) {
            //example to call the users API
            callYammerAPI(yammerConfig.yammerUrl, tokenResponse.accessToken, "users/current.json");
        }).catch(function (error) {
            console.log(error);
            // Upon acquireTokenSilent failure (due to consent or interaction or login required ONLY)
            // Call acquireTokenPopup(popup window) 
            if (requiresInteraction(error.errorCode)) {
                myMSALObj.acquireTokenPopup(requestObj).then(function (tokenResponse) {
                    callYammerAPI(yammerConfig.yammerUrl, tokenResponse.accessToken, "users.json");
                }).catch(function (error) {
                    console.log(error);
                });
            }
        });
    }

function callYammerAPI(baseUrl, accessToken, endpoint) {
        var request = new Request(baseUrl + endpoint, {
            headers: new Headers({
                'Authorization': 'Bearer ' + accessToken //specify accessToken as the authorization header. 
            })
        });

        //make the API request.
        fetch(request)
        .then(response => response.json())
        .then(function(data) {
            // data is your response. 
            // console.log(data); //response
            document.getElementById('json').innerHTML = document.getElementById('json').innerHTML + endpoint + " succeded.\n"
        }).catch(function(err){
            document.getElementById('json').innerHTML = document.getElementById('json').innerHTML + endpoint + " returned error.\n"
        });
    }
```


## Community Help and Support

We use [Stack Overflow](http://stackoverflow.com/questions/tagged/azure-active-directory) with the community to provide support. We highly recommend you ask your questions on Stack Overflow first and browse existing issues to see if someone has asked your question before.

Copyright (c) Microsoft Corporation.  All rights reserved. Licensed under the MIT License (the "License");
