# Office Auth Namespace (JavaScript API for Office)

The Office Auth namespace, Office.context.auth, provides a method that allows the Office host to obtain and access the add-in token. Indirectly, enable the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getAccessTokenAsync](office.context.auth.getAccessTokenAsync.md)|void|Calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application. Allows add-ins to identify users. Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).|

## Examples
For examples, see:
- [getAccessTokenAsync method](office.context.auth.getAccessTokenAsync.md)
- [Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)
- [Office Add-in ASP.NET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)

## More resources
[Enable single sign-on for Office Add-ins](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins)
