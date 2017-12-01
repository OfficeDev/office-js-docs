# RoamingSettings

The settings created by using the methods of the `RoamingSettings` object are saved per add-in and per user. That is, they are available only to the add-in that created them, and only from the user's mail box in which they are saved.

> While the Outlook Add-in API limits access to these settings to only the add-in that created them, these settings should not be considered secure storage. They can be accessed by Exchange Web Services or Extended MAPI. They should not be used to store sensitive information such as user credentials or security tokens.

The name of a setting is a String, while the value can be a String, Number, Boolean, null, Object, or Array.

The `RoamingSettings` object is accessible via the [`roamingSettings`](Office.context.md#roamingsettings-roamingsettings) property in the `Office.context` namespace.

> **Important:** The `RoamingSettings` object is initialized from the persisted storage only when the add-in is first loaded. For task panes, this means that it is only initialized when the task pane first opens. If the task pane navigates to another page or reloads the current page, the in-memory object is reset to its initial values, even if your add-in has persisted changes. The persisted changes will not be available until the task pane is closed and reopened.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Applicable Outlook mode| Compose or read|

##### Members and methods

| Member | Type |
|--------|------|
| [get](#getname--nullable-stringnumberbooleanobjectarray) | Method |
| [remove](#removename) | Method |
| [saveAsync](#saveasynccallback) | Method |
| [set](#setname-value) | Method |

### Example

```
// Get the current value of the 'myKey' setting
var value = Office.context.roamingSettings.get('myKey');
// Update the value of the 'myKey' setting
Office.context.roamingSettings.set('myKey', 'Hello World!');
// Persist the change
Office.context.roamingSettings.saveAsync();
```

### Methods

####  get(name) → (nullable) {String|Number|Boolean|Object|Array}

Retrieves the specified setting.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`name`| String|The case-sensitive name of the setting to retrieve.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Applicable Outlook mode| Compose or read|

##### Returns:

<dl class="param-type">

<dt>Type</dt>

<dd>String | Number | Boolean | Object | Array</dd>

</dl>

####  remove(name)

Removes the specified setting.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`name`| String|The case-sensitive name of the setting to remove.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Applicable Outlook mode| Compose or read|
####  saveAsync([callback])

Saves the settings.

Any settings previously saved by an add-in are loaded when it is initialized, so during the lifetime of the session you can just use the [`set`](RoamingSettings.md#setname-value) and [`get`](RoamingSettings.md#getname--nullable-stringnumberbooleanobjectarray) methods to work with the in-memory copy of the settings property bag. When you want to persist the settings so that they are available the next time the add-in is used, use the `saveAsync` method.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Applicable Outlook mode| Compose or read|
####  set(name, value)

Sets or creates the specified setting.

The set method creates a new setting of the specified name if it does not already exist, or sets an existing setting of the specified name. The value is stored in the document as the serialized JSON representation of its data type.

A maximum of 2MB is available for the settings of each add-in, and each individual setting is limited to 32KB.

Any changes made to settings using the `set` function will not be saved to the server until the [`saveAsync`](RoamingSettings.md#saveasynccallback) function is called.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`name`| String|The case-sensitive name of the setting to set or create.|
|`value`| String &#124; Number &#124; Boolean &#124; Object &#124; Array|The value to be stored.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Applicable Outlook mode| Compose or read|
