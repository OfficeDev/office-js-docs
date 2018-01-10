# Outlook add-in API Preview requirement set

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.

> **Note**: This documentation is for a **preview** [requirement set](tutorial-api-requirement-sets.html). This requirement set is not fully implemented yet, and clients will not accurately report support for it. You should not specify this requirement set in your add-in manifest. Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.

The Preview Requirement set includes all of the features of [Requirement set 1.6](../1.6/index.md). 

## Features in preview

The following features are in preview.

- [Event.completed](https://dev.office.com/reference/add-ins/outlook/preview/Event?product=outlook&version=preview#completedoptions) - A new optional parameter `options`, which is a dictionary with one valid value `allowEvent`. This value is used to cancel execution of an event.
- [Office.context.mailbox.item.getInitializationContextAsync](https://dev.office.com/reference/add-ins/outlook/preview/Office.context.mailbox.item?product=outlook&version=preview#getinitializationcontextasync) - Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).
- [Office.context.auth.getAccessTokenAsync](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync?product=outlook) - Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.

## Additional resources

- [Outlook add-ins](https://docs.microsoft.com/en-us/outlook/add-ins/)
- [Outlook add-in code samples](https://developer.microsoft.com/en-us/outlook/gallery/?filterBy=Outlook,Samples,Office add-ins)
- [Get started](https://docs.microsoft.com/en-us/outlook/add-ins/quick-start)
