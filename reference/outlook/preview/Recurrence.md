

# Recurrence

The `recurrence` object provides methods to get and set the recurrence pattern of appointments but only to get the recurrence pattern of meeting requests. It will have a dictionary with the following keys: `seriesTime`, `recurrenceType`, `recurrenceProperties`, and `recurrenceTimeZone` (optional).

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

### Members

#### seriesTime :Object

This object represents the startTime, endTime, startDate, and endDate of the series. **This is not in UTC time.** It is set in the time zone specified by the recurrenceTimeZone value or defaulted to the item's time zone.

##### Type:

* Object

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### recurrenceType :[Office.MailboxEnums.RecurrenceType](Office.MailboxEnums.md#recurrencetype-string)

This is one of the following: daily, weekday, weekly, monthly, yearly.

##### Type:

* [Office.MailboxEnums.RecurrenceType](Office.MailboxEnums.md#recurrencetype-string)

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### recurrenceProperties :Object

This is one of the following: interval, dayOfMonth, dayOfWeek, days, weekNumber, month, firstDayOfWeek

##### Type:

* interval|dayOfMonth|dayOfWeek|days|weekNumber|month|firstDayOfWeek

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### recurrenceTimeZone :[RecurrenceTimeZone](simple-types.md#recurrencetimezone)

Gets or sets the time zone of the recurrence.

##### Type:

* [RecurrenceTimeZone](simple-types.md#recurrencetimezone)

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

### Methods

#### getAsync([options], [callback])

Returns the current recurrence object.

This method returns the entire recurrence object.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`options`|Object|&lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`|Object|&lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`callback`|function|&lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object.|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

##### Examples

This example gets the recurrence object of an appointment or meeting request item.

```js
Office.context.mailbox.item.recurrence.getAsync(callback);
function callback(asyncResult){
	var context = asyncResult.context;
	var recurrence = asyncResult.value;
	var endDate = recurrence.seriesTime.getEndDate();
}
```

The following is an example of the results of the `getAsync` call.

```js
Recurrence = {"recurrenceType": "weekly","recurrenceProperties": {"interval": 2, "days": ["mon","thu","fri"], "firstDayOfWeek": "sun"}, "seriesTime": {seriesTimeObject}, "recurrenceTimeZone": {"name": "Pacific Standard Time", "offset": -480}}
```

####  setAsync(recurrencePattern, [options], [callback])

Sets the recurrence pattern of an appointment.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`recurrencePattern`|[Recurrence](recurrence.md)||A recurrence object.|
|`options`|Object|&lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`|Object|&lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`callback`|function|&lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object.|

##### Errors

|Error code|Description|
|------------|-------------|
|`InvalidEndTime`|The appointment end time is before the appointment start time.|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

> Note: `setAsync` should only be available for series items and not occurrence items.

##### Examples

The following example sets the recurrence pattern of an appointment series.

```js
var seriesTime= new seriesTime();
st.setStartDate(2017, 12, 1);   //required
st.setEndDate(2017, 12, 24);   //optional (if not ever set = null). Also can call st.setEndDate(null);
st.setStartTime(13, 30);       //required
st.setDuration(120);           //required

Office.context.mailbox.item.recurrence.setAsync(
	{	
		seriesTime,             //required
        recurrenceType,         //required
		recurrenceProperties,   //required
		recurrenceTimeZone      //optional
	}
	options,					//optional 
	callback					//optional
);
```
