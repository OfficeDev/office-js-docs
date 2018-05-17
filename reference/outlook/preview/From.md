

# From

Provides a method to get the from value of an appointment or message in an Outlook add-in.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)|ReadItem|
|Applicable Outlook mode|Compose|

### Methods

####  getAsync([options], callback)

Gets the from value of an appointment or message.

The `getAsync` method starts an asynchronous call to the Exchange server to get the from value of an appointment or message.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`callback`| function||When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object.

The from value of the item is provided as a string in the `asyncResult.value` property.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)|ReadItem|
|Applicable Outlook mode|Compose|