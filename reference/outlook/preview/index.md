# Outlook add-in API Preview requirement set

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.

> **Note**: This documentation is for a **preview** [requirement set](tutorial-api-requirement-sets.html). This requirement set is not fully implemented yet, and clients will not accurately report support for it. You should not specify this requirement set in your add-in manifest. Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.

The Preview Requirement set includes all of the features of [Requirement set 1.6](../1.6/index.md). 

## Features in preview

The following features are in preview.

- [From](https://dev.office.com/reference/add-ins/outlook/preview/From?product=outlook&version=preview) - Added a new object that provides a method to get the from value in an Outlook add-in.
- [Recurrence](https://dev.office.com/reference/add-ins/outlook/preview/Recurrence?product=outlook&version=preview) - Added a new object that provides methods to get and set the recurrence pattern of appointments but only get the recurrence pattern of messages that are meeting requests.
- [SeriesTime](https://dev.office.com/reference/add-ins/outlook/preview/SeriesTime?product=outlook&version=preview) - Added a new object that provides methods to get and set the dates and times of appointments in a recurring series and to get the dates and times of meeting requests in a recurring series.
- [Event.completed](https://dev.office.com/reference/add-ins/outlook/preview/Event?product=outlook&version=preview#completedoptions) - A new optional parameter `options`, which is a dictionary with one valid value `allowEvent`. This value is used to cancel execution of an event.
- [Office.context.mailbox.item.addHandlerAsync](https://dev.office.com/reference/add-ins/outlook/preview/Office.context.mailbox.item?product=outlook&version=preview#addhandlerasynceventtype-handler-options-callback) - Added a new method that adds an event handler for a supported event.
- [Office.context.mailbox.item.from](https://dev.office.com/reference/add-ins/outlook/preview/Office.context.mailbox.item?product=outlook&version=preview#from-emailaddressdetailsfrom) - Modified to get the from value in Compose mode.
- [Office.context.mailbox.item.getInitializationContextAsync](https://dev.office.com/reference/add-ins/outlook/preview/Office.context.mailbox.item?product=outlook&version=preview#getinitializationcontextasyncoptions-callback) - Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).
- [Office.context.mailbox.item.recurrence](https://dev.office.com/reference/add-ins/outlook/preview/Office.context.mailbox.item?product=outlook&version=preview#nullable-recurrence-recurrence) - Added a new property that gets or sets an object which provides methods to manage the recurrence pattern of an appointment item. This property can also be used to get the recurrence pattern of a meeting request item.
- [Office.context.mailbox.item.removeHandlerAsync](https://dev.office.com/reference/add-ins/outlook/preview/Office.context.mailbox.item?product=outlook&version=preview#removehandlerasynceventtype-handler-options-callback) - Added a new method that removes an event handler.
- [Office.context.mailbox.item.seriesId](https://dev.office.com/reference/add-ins/outlook/preview/Office.context.mailbox.item?product=outlook&version=preview#nullable-seriesid-string) - Added a new property that gets the id of the series an occurrence belongs to.
- [Office.context.auth.getAccessTokenAsync](https://dev.office.com/reference/add-ins/shared/Office.context.auth.getAccessTokenAsync?product=outlook) - Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.
- [Office.MailboxEnums.Days](https://dev.office.com/reference/add-ins/outlook/preview/Office.MailboxEnums?product=outlook&version=preview#days-string) - Added a new enum that specifies the day of week or type of day. 
- [Office.MailboxEnums.Month](https://dev.office.com/reference/add-ins/outlook/preview/Office.MailboxEnums?product=outlook&version=preview#month-string) - Added a new enum that specifies the month.
- [Office.MailboxEnums.RecurrenceTimeZone](https://dev.office.com/reference/add-ins/outlook/preview/Office.MailboxEnums?product=outlook&version=preview#recurrencetimezone-string) - Added a new enum that specifies the time zone applied to the recurrence.
- [Office.MailboxEnums.RecurrenceType](https://dev.office.com/reference/add-ins/outlook/preview/Office.MailboxEnums?product=outlook&version=preview#recurrencetype-string) - Added a new enum that specifies the type of recurrence. 
- [Office.MailboxEnums.WeekNumber](https://dev.office.com/reference/add-ins/outlook/preview/Office.MailboxEnums?product=outlook&version=preview#weeknumber-string) - Added a new enum that specifies the week of the month.

## See also

- [Outlook add-ins](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](https://docs.microsoft.com/outlook/add-ins/quick-start)
