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

## Requirements

|Host|Introduced in|Last changed in|
|:---------------|:--------|:----------|
|Word, Excel, PowerPoint|1.1|1.1|
|Outlook|Mailbox 1.4|Mailbox 1.4|

This method is available in the SSOAddin [requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md). To specify the SSOAddin requirement set, use the following in your manifest.

```xml
 <Requirements> 
   <Sets DefaultMinVersion="1.1"> 
     <Set Name="SSOAddin"/> 
   </Sets> 
 </Requirements> 

```

To detect this API at runtime, use the following code.

```js
 if (Office.context.requirements.isSetSupported('SSOAddin', 1.1)) 
 	{  
    	 // Use Office UI methods; 
 	} 
 else 
	 { 
	     // Alternate path 
	 } 
```

## Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|callback|object|Accepts a callback method to handle the identity objects.|
|forceAction|string|Optional. Accepts a single string to prompt the Office for user input.|

## Callback method
When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **requestAuthToken** method, you can use the properties of the **AsyncResult** object to return [add-in identity objects](#Identity-Objects).

### forceAction options
The following actions are available during authentication.

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|**forceAddinConsent**|string|Optional. Causes Office to display the add-in consent experience. Useful if the add-in's Azure permissions have changed or if the user's consent has been revoked.|
|**forceAddAccount**|object|Optional. Prompts the user to add (or to switch if already added) his or her Office account.|

## Remarks

This API requires a single sign-on configuration that bridges the add-in to an Azure application. Office users sign-in with Organizational Accounts and Microsoft Accounts. Microsoft Azure returns Tokens intended for both user account types to access resources in the Microsoft Graph.

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

## Identity Objects
requestAuthTokenAsync returns an array of objects via the asyncResult.value with the following characteristics:

|**Property**|**Use to**|
|:-----|:-----|
|DirectoryType|Describes whether the attached Office account is an Organizational Account or a Microsoft Account.|
|AccessToken|The Access Token, issued to your Add-in's Web service, returned from Azure on the user's behalf.|

## Support history
Page created.

****

|**Version**|**Changes**|
|:-----|:-----|
|1.3|Introduced|
