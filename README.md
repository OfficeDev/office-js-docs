## Office Add-in Single Sign-On

This repository describes the design and steps for enabling single sign-on for Office Add-ins. 

The main method for acquring the access token is described [here](https://github.com/OfficeDev/office-js-docs/blob/Addin_SSO_OpenSpec/reference/shared/office.context.auth.getAccessTokenAsync.md)

For some time, users have been able to sign into Office (online, mobile, and desktop platforms) with either their personal (Microsoft Account) credentials or their school or work (Office 365) credentials. 
 
Your add-in can now take advantage of this sign-in process to
* authorize the user to your add-in
* authorize the add-in to access Microsoft Graph 
without requiring the user to sign-in a second time. 
For users, this makes running your add-in a smooth experience, usually involving at most a one-time consent screen. For developers, this means that your add-in can authenticate users  and gain authorized access to the user’s   data via Microsoft Graph with credentials that the Office application itself has already gathered.

NOTE: FOR THIS DEVELOPER PREVIEW, SINGLE SIGN-ON IS INITIALLY SUPPORTED ONLY FOR WORK OR SCHOOL (OFFICE 365) ACCOUNTS  AND ONLY FOR DESKTOP VERSIONS OF OFFICE. 

### Add-in Architecture
In addition to hosting the pages and JavaScript of the web application, the add-in must also host, at the same fully qualified domain , one or more Web APIs which will obtain an access token to Microsoft Graph and make requests to it.

> Fully qualified domain name - A name that unambiguously identifies, or fully qualifies, a computer within the entire DNS hierarchy for example, server1.widgets.microsoft.com

The add-in manifest contains markup that specifies how the add-in is registered in the Azure AD V2 endpoint, and it specifies any permissions to the Microsoft Graph that the add-in needs.
How it Works at Runtime
The diagram below shows how the SSO process works. 
 
1.	JavaScript in the add-in calls a new Office.js API getAccessTokenAsync. This tells the Office host application to obtain an access token to the add-in. (Hereafter, this is called the “add-in token”.)
2.	[Occurs only if needed] If the user is not signed in, the Office host application opens a popup for the user to sign in. 
3.	[Occurs only if needed] If this is the first time the current user has used your add-in, s/he is prompted to consent. 
4.	The Office host application requests the add-in token from the Azure AD V2 endpoint for the current user.
5.	Azure AD sends the add-in token to the Office host application.
6.	The Office host application sends the add-in token to the add-in as part of the result object returned by the call of getAccessTokenAsync.
7.	JavaScript in the add-in makes an HTTP Request to a Web API that is hosted at the same fully-qualified domain as the add-in, and it includes the add-in token as authorization proof .  
8.	Server-side code validates the incoming add-in token.
9.	Server-side code uses the “on behalf of” flow (defined at OAuth2 Token Exchange and this Azure Scenario guide) to obtain an access token for Microsoft Graph (hereafter, the “MSG token”) in exchange for the add-in token. 
10.	AAD returns the MSG token (and a refresh token, if the add-in requests offline_access permission) to the add-in.
11.	Server-side code caches the token(s).
12.	Server-side code makes Requests to Microsoft Graph and includes the MSG token.
13.	Microsoft Graph returns data to the add-in, which can pass it on to the add-in’s UI. 
14.	[Occurs as needed] When the MSG token expires, the server-side code can use its refresh token to obtain a new MSG token.

### Development Tasks
These are the major steps to creating an Office Add-in that uses Single Sign-on described in a language-agnostic, framework-agnostic way. A detailed walkthrough for an ASP.NET MVC add-in is being written.
#### Step 1 – Create the Service Application
Register the add-in at the Azure V2 Endpoint’s registration portal: apps.dev.microsoft.com. This is a 5 – 10 minute process that includes:
* Obtaining a client ID and secret for the add-in.
* Specifying the permissions that your add-in needs to Microsoft Graph.
* Preauthorizing the Office host application (Client Id d3590ed6-52b3-4102-aeff-aad2292ab01c) to the add-in with the default permission “access_as_user”.
#### Step 2 – Configure the add-in
Add new markup to the add-in manifest:
* WebApplicationID is the client ID of the add-in.
* WebApplicationResource is the URL of the add-in.
* WebApplicationScopes specifies the permissions that the Office host needs to the add-in and that the add-in needs to MS Graph. In general, you’ll always want User.Read, but you can request more access (e.g. Mail.Read or offline_access).
#### Step 3 – Code the client-side
Add JavaScript to the add-in to:
*	Call Office.context.auth.getAccessTokenAsync(myTokenHandler).
*	Create a handler that passes the add-in token to the add-in’s server-side code. For example:

~~~~ javascript 
function mytokenHandler(asyncResult) { 
    // passes asyncResult.value (which has the add-in access token)
    // to the add-in’s Web API as a Bearer type Authorization header.
}
~~~~

#### Step 4 – Code the server-side
Create one or more Web API methods that get MS Graph data. Depending on your language and framework, there may be libraries that will greatly simplify the code you have to write. Your server-side code needs to do the following:
*	Validate the add-in token that is received from the token handler you created in Step 3.
*	Initiate the “on behalf of” flow with a call to the Azure AD V2 endpoint that includes the add-in access token, some metadata about the user, and the credentials of the add-in (it’s ID and secret). 
*	Cache the returned MSG token.
*	Get data from MS Graph using the MSG token.

_**Note**: this feature is still under design and review phase and hence not yet available as part of the product. The final design is subject to change. Once the feature is made available, the final specification will be published as part of the master repository._

## Give us your feedback

Your feedback is important to us. 
* To let us know about any questions or issues you find in the docs, [submit an issue](https://github.com/OfficeDev/office-js-docs/issues) in this repository. Make sure you state the version+build number of the client you are using, and provide repro steps, console output, and error messages, as appropriate. 
* We also encourage you to fork, make the fix, and do a pull request of your proposed changes. For details, see [Contribute to this documentation](Contributing.md). 
* To let us know about your programming experience, what you would like to see in future versions, code samples, and so on, enter your suggestions and ideas at [Office Developer Platform UserVoice](https://officespdev.uservoice.com/).

## Copyright

Copyright (c) 2016 Microsoft Corporation. All rights reserved.
