

# Recurrence

The `recurrence` object provides methods to get and set the recurrence pattern of appointments but only to get the recurrence pattern of meeting requests. It will have a dictionary with the following keys: `seriesTime`, `recurrenceType`, `recurrenceProperties`, and `recurrenceTimeZone` (optional).

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

##### States

|State|Editable?|Viewable?|
|---|---|---|
|Appointment Organizer - Compose Series|Yes (`setAsync`)|Yes (`getAsync`)|
|Appointment Organizer - Compose Instance|No (`setAsync` returns error)|Yes (`getAsync`)|
|Appointment Attendee - Read Series|No (`setAsync` not available)|Yes (`item.recurrence`)|
|Appointment Attendee - Read Instance|No (`setAsync` not available)|Yes (`item.recurrence`)|
|Meeting Request - Read Series|No (`setAsync` not available)|Yes (`item.recurrence`)|
|Meeting Request - Read Instance|No (`setAsync` not available)|Yes (`item.recurrence`)|

### Members

#### recurrenceProperties :[RecurrenceProperties](simple-types.md#recurrenceproperties)

Gets or sets the properties of the recurrence.

##### Type:

* [RecurrenceProperties](simple-types.md#recurrenceproperties)

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

|[Recurrence type](Office.MailboxEnums.md#recurrencetype-string)|Applicable [recurrence properties](simple-types.md#recurrenceproperties)|Description|
|---|---|---|
|`daily`|-`interval`|An appointment occurs every **2** *{`interval`}* days.|
|`weekday`||An appointment occurs every weekday.|
|`monthly`|-`interval`<br>-`dayOfMonth`<br>-`dayOfWeek`<br>-`weekNumber`|An appointment occurs on day **5** *{`dayOfMonth`}* every **4** *{`interval`}* months.<br>An appointment occurs on the **third** *{`weekNumber`}* **thu** *{`dayOfWeek`}* every **2** *{`interval`}* months.|
|`weekly`|-`interval`<br>-`days`<br>-`firstDayOfWeek` (optional)|An appointment occurs on **tue** and **thu** *{`days`}* every **2** *{`interval`}* weeks.|
|`yearly`|-`interval`<br>-`dayOfMonth`<br>-`dayOfWeek`<br>-`weekNumber`<br>-`month`|An appointment occurs on day **7** *{`dayOfMonth`}* of **sep** *{`month`}* every **4** *{`interval`}* years.<br>An appointment occurs on the **first** *{`weekNumber`}* **Thu** *{`dayOfWeek`}* of **sep** *{`month`}* every **2** *{`interval`}* years.|

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
var seriesTimeObject = new seriesTime(); 
seriesTimeObject.setStartDate(2017,11,2);  
seriesTimeObject.setEndDate(2017,12,2); 
seriesTimeObject.setStartTime(10,30); 
SeriesTimeObject.setDuration(30);

var pattern = {"seriesTime": seriesTimeObject, "type": "Weekly", "properties": {"interval": ", "days": ["Tue", "Thu"], "firstDayOfWeek": "Sun"}, "recurrenceTimeZone": {"name": "Pacific Standard Time"}}; 

Office.context.mailbox.item.recurrence.setAsync(pattern, options, callback);
 
//Result: This created a recurring event from November 2, 2017 to December 2, 2017 at 10:30 A.M. to 11 A.M. PST every Tuesday and Thursday.
```

### Errors

|Error|Error message|Description|
|---|---|---|
|Sys.ArgumentTypeException|Object of type {the type of the incorrect parameter} cannot be converted to type '{the type of the correct parameter}'. Parameter name: {parameter name}<br><br>Example: Object of type 'String' cannot be converted to type 'Date'. Parameter name: startDate|This occurs when the wrong argument type is given. For example, a string is given instead of a datetime object.|
|Invalid Value|The end date cannot be before the start date.|When the user specifies an end date that is before the start date.|
|Permission Denied|A recurrence pattern cannot be set at an instance level.|When an add-in calls recurrence.setAsync on an instance item.|
|Sys.ArgumentException|Value does not fall within the expected range. Parameter name: {parameter name}.<br><br>Example: Value does not fall within the expected range. Parameter name: dayOfWeek|When the add-in uses an incorrect value as a parameter for an enum like dayOfWeek, weekNumber, type, days, or month.|
|Invalid Value|The recurrence pattern is invalid. Double check that the specified recurrence properties aligns with the recurrence type.|When the add-in uses an incorrect recurrence pattern. For example, they use Daily but specify Weekday or the user specifies Monthly then uses dayOfMonth and dayOfWeek.|
|Missing Parameter|The recurrence object could not be created because some parameter values are missing.|When the add-in tries to set the recurrence pattern but does not include necessary parameters. For example, the recurrence type is daily but the Interval parameter is missing.|
|Invalid Value|The specified time zone is not supported.|When the add-in tries to set the time zone of a recurrence pattern but enters a string that does not match one of the supported time zones.|
|Invalid Value|The recurrence pattern was set by the user using an alternate calendar that is not supported.|When the user uses an alternate calendar to set recurrence such as a lunar calendar.|