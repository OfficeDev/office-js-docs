

# Recipients

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose|

### Methods

####  addAsync(recipients, [options], [callback])

Adds a recipient list to the existing recipients for an appointment or message.

The `recipients` parameter can be an array of one of the following:

*   Strings containing SMTP email addresses
*   `EmailUser` objects
*   `EmailAddressDetails` objects

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`recipients`| Array.&lt;(String&#124;[EmailUser](simple-types.md#emailuser)&#124;[EmailAddressDetails](simple-types.md#emailaddressdetails))&gt;||The recipients to add to the recipients list.|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. <br/>If adding the recipients fails, the `asyncResult.error` property will contain an error code.|

##### Errors

| Error code | Description |
|------------|-------------|
| `NumberOfRecipientsExceeded` | The number of recipients exceeded 100 entries. |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Applicable Outlook mode| Compose|

##### Example

The following example creates an array of `EmailUser` objects and adds them to the To recipients of the message.

```
var newRecipients = [
  {
    "displayName": "Allie Bellew",
    "emailAddress": "allieb@contoso.com"
  },
  {
    "displayName": "Alex Darrow",
    "emailAddress": "alexd@contoso.com"
  }
];

Office.context.mailbox.item.to.addAsync(newRecipients, function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Recipients added");
  }
});
```

####  getAsync([options], callback)

Gets a recipient list for an appointment or message.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`callback`| function||When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object.

When the call completes, the `asyncResult.value` property will contain an array of [`EmailAddressDetails`](simple-types.md#emailaddressdetails) objects.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose|

##### Example

The following example gets the optional attendees of a meeting.

```js
Office.context.mailbox.item.optionalAttendees.getAsync(function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    var msg = "";
    result.value.forEach(function(recip, index) {
      msg = msg + recip.displayName + " (" + recip.emailAddress + ");";
    });
    showMessage(msg);
  }
});
```

####  setAsync(recipients, [options], callback)

Sets a recipient list for an appointment or message.

The `setAsync` method overwrites the current recipient list.

The `recipients` parameter can be an array of one of the following:

*   Strings containing SMTP email addresses
*   `EmailUser` objects
*   `EmailAddressDetails` objects

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`recipients`| Array.&lt;(String&#124;[EmailUser](simple-types.md#emailuser)&#124;[EmailAddressDetails](simple-types.md#emailaddressdetails))&gt;||The recipients to add to the recipients list.|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`callback`| function||When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. <br/>If setting the recipients fails the `asyncResult.error` property will contain a code that indicates any error that occurred while adding the data.|

##### Errors

| Error code | Description |
|------------|-------------|
| `NumberOfRecipientsExceeded` | The number of recipients exceeded 100 entries. |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Applicable Outlook mode| Compose|

##### Example

The following example creates an array of `EmailUser` objects and replaces the CC recipients of the message with the array.

```
var newRecipients = [
  {
    "displayName": "Allie Bellew",
    "emailAddress": "allieb@contoso.com"
  },
  {
    "displayName": "Alex Darrow",
    "emailAddress": "alexd@contoso.com"
  }
];

Office.context.mailbox.item.cc.setAsync(newRecipients, function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Recipients overwritten");
  }
});
```
