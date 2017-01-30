# requestAuthTokenAsync method
Requests an authentication Token for the user currently signed into Office.

 **Important:** This API currently works only in Excel, Outlook, PowerPoint, Word, and OneNote in [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview).

|||
|:-----|:-----|
|**Hosts:**|Excel, Outlook, PowerPoint, Word, OneNote|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|SSOAddIn|
|**Added in**|1.3|

```
var swapToken = Office.context.auth.requestAuthTokenAsync(handleAuthToken);
```

## Return value

An authentication Token.


## Remarks

This API requires a single sign-on configuration that bridges the add-in to an Azure application. Office users sign-in with Organizational Accounts and Microsoft Accounts. Microsoft Azure returns Tokens intended for both user accounts types to access resources in the Microsoft Graph.


## Support details

A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).

**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|**OWA for Devices**|
|:-----|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|Y|
|**Outlook**|Y|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|Y|
|**Word**|Y|Y|Y|Y|
|**OneNote**|Y|Y|Y|Y|

|||
|:-----|:-----|
|**Minimum permission level**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane, Outlook|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.3|Introduced|
