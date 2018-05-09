

# Recurrence

The `recurrence` object provides methods to get and set the recurrence pattern of appointments but only to get the recurrence pattern of meeting requests. It will have a dictionary with the following keys: `seriesTime`, `recurrenceType`, `recurrenceProperties`, and `recurrenceTimeZone` (optional).

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

### Members

#### recurrenceProperties :Object

Gets or sets the properties of the recurrence.

##### Type:

* Object

##### Properties:

|Name|Type|Description|
|---|---|---|
|`interval`|Integer|Represents the period between recurrences of the same meeting or appointment series.|
|`dayOfMonth`|Integer|Represents the day of the month.|
|`dayOfWeek`|[Office.MailboxEnums.Days](Office.MailboxEnums.md#days-string)|Represents the day of the week or type of day, for example, weekend day vs weekday.|
|`days`|Array of [Office.MailboxEnums.Days](Office.MailboxEnums.md#days-string)|Represents the set of days for this recurrence. Valid values are: `Mon`, `Tue`, `Wed`, `Thu`, `Fri`, `Sat`, and `Sun`.|
|`weekNumber`|[Office.MailboxEnums.WeekNumber](Office.MailboxEnums.md#weeknumber-string)|Represents the number of the week in the selected month e.g. `first` for first week of the month.|
|`month`|[Office.MailboxEnums.Month](Office.MailboxEnums.md#month-string)|Represent the month.|
|`firstDayOfWeek`|[Office.MailboxEnums.Days](Office.MailboxEnums.md#days-string)|Represents the first day of the week you choose otherwise the default is the value in the current user's settings. Valid values are: `Mon`, `Tue`, `Wed`, `Thu`, `Fri`, `Sat`, and `Sun`.|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### recurrenceTimeZone :[RecurrenceTimeZone](simple-types.md#recurrencetimezone)

Gets or sets the time zone of the recurrence.

##### Type:

* [RecurrenceTimeZone](simple-types.md#recurrencetimezone)

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### recurrenceType :[Office.MailboxEnums.RecurrenceType](Office.MailboxEnums.md#recurrencetype-string)

Gets or set the type of the recurrence.

##### Type:

* [Office.MailboxEnums.RecurrenceType](Office.MailboxEnums.md#recurrencetype-string)

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### seriesTime :[seriesTime](seriestime.md)

This object enables you to manage the start and end dates and times of the series and of each recurrence. **This is not in UTC time.** It is set in the time zone specified by the [recurrenceTimeZone](#recurrencetimezone-recurrencetimezone) value or defaulted to the item's time zone.

##### Type:

* [seriesTime](seriestime.md)

##### Requirements

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
st.setEndDate(2017, 12, 24);   //optional (if not ever set, equals null). Also can call st.setEndDate(null);
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
