

# Recurrence

The `recurrence` object provides methods to get and set the recurrence pattern of appointments but only get the recurrence pattern of meeting requests. It will have a dictionary with the following keys: `seriesTime`, `recurrenceType`, `recurrenceProperties`, and `recurrenceTimeZone` (optional).

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

Gets or sets the properties of the recurring appointment series.

##### Type:

* [RecurrenceProperties](simple-types.md#recurrenceproperties)

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### recurrenceTimeZone :[RecurrenceTimeZone](simple-types.md#recurrencetimezone)

Gets or sets the time zone of the recurring appointment series

##### Type:

* [RecurrenceTimeZone](simple-types.md#recurrencetimezone)

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### recurrenceType :[Office.MailboxEnums.RecurrenceType](Office.MailboxEnums.md#recurrencetype-string)

Gets or sets the type of the recurring appointment series.

|[Recurrence type](Office.MailboxEnums.md#recurrencetype-string)|Applicable [recurrence properties](simple-types.md#recurrenceproperties)|Description|
|---|---|---|
|`daily`|-`interval`|An appointment occurs every **2** *{`interval`}* days.|
|`weekday`||An appointment occurs every weekday.|
|`monthly`|-`interval`<br>-`dayOfMonth`<br>-`dayOfWeek`<br>-`weekNumber`|An appointment occurs on day **5** *{`dayOfMonth`}* every **4** *{`interval`}* months.<br>An appointment occurs on the **Third** *{`weekNumber`}* **Thu** *{`dayOfWeek`}* every **2** *{`interval`}* months.|
|`weekly`|-`interval`<br>-`days`<br>-`firstDayOfWeek` (optional)|An appointment occurs on **Tue** and **Thu** *{`days`}* every **2** *{`interval`}* weeks.|
|`yearly`|-`interval`<br>-`dayOfMonth`<br>-`dayOfWeek`<br>-`weekNumber`<br>-`month`|An appointment occurs on day **7** *{`dayOfMonth`}* of **Sep** *{`month`}* every **4** *{`interval`}* years.<br>An appointment occurs on the **First** *{`weekNumber`}* **Thu** *{`dayOfWeek`}* of **Sep** *{`month`}* every **2** *{`interval`}* years.|

##### Type:

* [Office.MailboxEnums.RecurrenceType](Office.MailboxEnums.md#recurrencetype-string)

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### seriesTime :[seriesTime](seriestime.md)

This object enables you to manage the start and end dates of the recurring appointment series and the usual start and end times of instances. **This object is not in UTC time.** Instead, it is set in the time zone specified by the [recurrenceTimeZone](#recurrencetimezone-recurrencetimezone) value or defaulted to the item's time zone.

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

Returns the current recurrence object of an appointment series.

This method returns the entire recurrence object for the appointment series.

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

Sets the recurrence pattern of an appointment series.

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
|`InvalidEndTime`|The appointment end time is before its start time.|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

> Note: `setAsync` should only be available for series items and not instance items.

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
