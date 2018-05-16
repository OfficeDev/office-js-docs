

# SeriesTime

The `seriesTime` object provides methods to get and set the dates and times of meetings or appointments in a series and the start and end of the series.

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

### Methods

#### getDuration() → {Integer}

Gets the duration of a recurrence pattern in minutes.

##### Examples

This example gets the end date of a recurrence.

```js
Office.context.mailbox.item.recurrence.getAsync(callback);
function callback(asyncResult){
	var context = asyncResult.context;
	var recurrence = asyncResult.value;
	var duration = recurrence.seriesTime.getDuration();
}
```

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### getEndDate() → {String}

Gets the end date of a recurrence pattern in the following [ISO 8601](https://www.iso.org/iso-8601-date-and-time-format.html) date format:
"YYYY-MM-DD"

##### Examples

This example gets the end date of a recurrence.

```js
Office.context.mailbox.item.recurrence.getAsync(callback);
function callback(asyncResult){
	var context = asyncResult.context;
	var recurrence = asyncResult.value;
	var endDate = recurrence.seriesTime.getEndDate();
}
```

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### getEndTime() → {String}

Gets the start time of the meeting/appointment of a recurrence pattern in whichever time zone that the user/add-in set the recurrence pattern in the following ISO 8601 format:
"THH:mm:ss:mmm"

##### Examples

This example gets the end time of a recurrence.

```js
Office.context.mailbox.item.recurrence.getAsync(callback);
function callback(asyncResult){
	var context = asyncResult.context;
	var recurrence = asyncResult.value;
	var endDate = recurrence.seriesTime.getEndTime();
}
```

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### getStartDate() → {String}

Gets the start date of a recurrence pattern in the following [ISO 8601](https://www.iso.org/iso-8601-date-and-time-format.html) date format:
"YYYY-MM-DD"

##### Examples

This example gets the start date of a recurrence.

```js
Office.context.mailbox.item.recurrence.getAsync(callback);
function callback(asyncResult){
	var context = asyncResult.context;
	var recurrence = asyncResult.value;
	var endDate = recurrence.seriesTime.getStartDate();
}
```

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### getStartTime() → {String}

Gets the start time of the meeting/appointment of a recurrence pattern in whichever time zone that the user/add-in set the recurrence pattern in the following ISO 8601 format:
"THH:mm:ss:mmm"

##### Examples

This example gets the start time of a recurrence.

```js
Office.context.mailbox.item.recurrence.getAsync(callback);
function callback(asyncResult){
	var context = asyncResult.context;
	var recurrence = asyncResult.value;
	var endDate = recurrence.seriesTime.getStartTime();
}
```

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### setDuration(int minutes)

Sets the duration of a meeting/appointment of a recurrence pattern. This should also change the end time of the recurrence pattern.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`minutes`|Integer||The length of the meeting or appointment in minutes.|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the duration of a recurrence to 2 hours.

```js
var seriesTimeObject = new seriesTime()
seriesTimeObject.setDuration(120);
```

#### setEndDate(int year, int month, int day)

Sets the end date of a recurrence pattern.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`year`|Integer||The year portion of the send date.|
|`month`|Integer||The month portion of the end date. Valid range is 0-11 where 0 represents the 1st month and 11 represents the 12th month.|
|`day`|Integer||The day portion of the end date.|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the end date of a recurrence to November 2, 2017.

```js
var seriesTimeObject = new seriesTime()
seriesTimeObject.setEndDate(2017, 10, 2);
```

##### Errors

|Error code|Description|
|------------|-------------|
|`Invalid date format`||

#### setEndDate(String date)

Sets the end date of a recurrence pattern.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`date`|String||End date of the recurrence represented in [ISO 8601](https://www.iso.org/iso-8601-date-and-time-format.html) date standard, that is, "YYYY-MM-DD".|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the end date of a recurrence to November 2, 2017.

```js
var seriesTimeObject = new seriesTime()
seriesTimeObject.setEndDate("2017-11-02");
```

##### Errors

|Error code|Description|
|------------|-------------|
|`Invalid date format`|The date is is not in an acceptable format.|

#### setStartTime(int hours, int minutes)

Sets the start time of the meeting/appointment of a recurrence pattern in whichever time zone the recurrence pattern is set. Default zone is the item's time zone.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`hours`|Integer||The `hour` portion of the start time. Valid range: 0-24.|
|`minutes`|Integer||The `minute` portion of the start time. Valid range: 0-59.|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the start time of a recurrence to 1:30 pm.

```js
var seriesTimeObject = new seriesTime()
seriesTimeObject.setStartTime(13, 30);
```

This example sets the start time of a recurrence to 11:30 am.

```js
var seriesTimeObject = new seriesTime()
seriesTimeObject.setStartTime(11, 30);
```

##### Errors

|Error code|Description|
|------------|-------------|
|`Invalid time format`|The time is is not in an acceptable format.|

#### setStartTime(String time)

Sets the start time of the meeting/appointment of a recurrence pattern in whichever time zone the recurrence pattern is set in. Default value is the item's time zone.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`time`|String||Start time of the recurrence represented in standard datetime string format, that is, "THH:mm:ss:mmm".|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the start time of a recurrence to 11:30 pm.

```js
var seriesTimeObject = new seriesTime()
seriesTimeObject.setStartTime("T23:30:00");
```

##### Errors

|Error code|Description|
|------------|-------------|
|`Invalid time format`|The time is is not in an acceptable format.|

#### setStartDate(int year, int month, int day)

Sets the start date of a recurrence pattern.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`year`|Integer||The year portion of the start date.|
|`month`|Integer||The month portion of the start date. Valid range is 0-11 where 0 represents the 1st month and 11 represents the 12th month.|
|`day`|Integer||The day portion of the start date.|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the start date of a recurrence to November 2, 2017.

```js
var seriesTimeObject = new seriesTime()
seriesTimeObject.setStartDate(2017, 10, 2);
```

##### Errors

|Error code|Description|
|------------|-------------|
|`Invalid date format`|The date is not in an acceptable format.|

#### setStartDate(String date)

Sets the start date of a recurrence pattern.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`date`|String||Start date of the recurrence in [ISO 8601](https://www.iso.org/iso-8601-date-and-time-format.html) date standard, that is, "YYYY-MM-DD".|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the start date of a recurrence to November 2, 2017.

```js
var seriesTimeObject = new seriesTime()
seriesTimeObject.setStartDate("2017-11-02");
```

##### Errors

|Error code|Description|
|------------|-------------|
|`Invalid date format`|The date is is not in an acceptable format.|

### Errors

|Error|Error message|Description|
|---|---|---|
|Sys.ArgumentException|Value does not fall within the expected range. Parameter name: {Parameter Name}.|**Methods**: setStartDate, setEndDate<br>The month is not a value in the range of 0-11 or the day is not a valid date (for example, Oct. 32 or Sept. 31). **Note**: A date in the past is okay.|
|Sys.ArgumentException|Value does not fall within the expected range. Parameter name: {Parameter Name}.|**Methods**: setStartTime<br>The hours are not in the range 0-24 or the minute values are not in the range 0-59.|
|Invalid Value|The end date cannot be before the start date.|When the user specifies an end date that is before the start date.|
|Missing Parameter|The seriesTime object could not be created because at least one required parameter value is missing.|The seriesTime object needs to have all of the following to be created: a start time, a duration, a start date, and an end date. If it does not have one of these values, this error will occur.|
|Missing Parameter|Missing parameter: {Parameter name}|For any of the setter methods, when a required parameter is missing from the method call.|