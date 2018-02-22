# Outlook add-in API requirement set 1.6

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.

## What's new in 1.6?

Requirement set 1.6 includes all of the features of [Requirement set 1.5](../1.5/index.md). It added the following features.

- Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.
- Added a new API to open a new message form.
- Added the ability for the add-in to determine the account type of the user's mailbox.

### Change log

- Added [Office.context.mailbox.item.getSelectedEntities](https://dev.office.com/reference/add-ins/outlook/1.6/Office.context.mailbox.item?product=outlook&version=v1.6#getselectedentities--entities): Adds a new function that gets the entities found in a highlighted match a user has selected. Highlighted matches apply to contextual add-ins.
- Added [Office.context.mailbox.item.getSelectedRegExMatches](https://dev.office.com/reference/add-ins/outlook/1.6/Office.context.mailbox.item?product=outlook&version=v1.6#getselectedregexmatches--object): Adds a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to contextual add-ins.
- Added [Office.context.mailbox.displayNewMessageForm](https://dev.office.com/reference/add-ins/outlook/1.6/Office.context.mailbox?product=outlook&version=v1.6#displaynewmessageformparameters): Adds a new function that opens a new message form.
- Added [Office.context.mailbox.userProfile.accountType](https://dev.office.com/reference/add-ins/outlook/1.6/Office.context.mailbox.userProfile?product=outlook&version=v1.6#accounttype-string): Adds a new member to the user profile that indicates the type of the user's account.

## Additional resources

- [Outlook add-ins](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Office add-ins)
- [Get started](https://docs.microsoft.com/outlook/add-ins/quick-start)
