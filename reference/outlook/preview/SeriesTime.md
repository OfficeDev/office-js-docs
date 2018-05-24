

# SeriesTime

The `seriesTime` object provides methods to get and set the dates and times of appointments in a recurring series and to get the dates and times of meeting requests in a recurring series.

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

### Methods

#### getDuration() → {Integer}

Gets the duration in minutes of a usual instance in a recurring appointment or meeting request series.

##### Examples

This example gets the duration of a usual instance in a recurring appointment or meeting request series.

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

This example gets the end date of a recurring appointment or meeting request series.

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

Gets the end time of a usual appointment or meeting request instance of a recurrence pattern in whichever time zone that the user/add-in set the recurrence pattern using the following ISO 8601 format:
"THH:mm:ss:mmm"

##### Examples

This example gets the end time of a usual instance in a recurring appointment or meeting request series.

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

This example gets the start date of a recurring appointment or meeting request series.

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

Gets the start time of a usual appointment or meeting request instance of a recurrence pattern in whichever time zone that the user/add-in set the recurrence pattern using the following ISO 8601 format:
"THH:mm:ss:mmm"

##### Examples

This example gets the start time of a usual instance in a recurring appointment or meeting request series.

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

#### setDuration(minutes)

Sets the duration of all appointments in a recurrence pattern. This will also change the end time of the recurrence pattern.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`minutes`|Integer||The length of the appointment in minutes.|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the duration of each appointment in a recurring series to 2 hours.

```js
var seriesTimeObject = new seriesTime()
seriesTimeObject.setDuration(120);
```

#### setEndDate(year, month, day)

Sets the end date of a recurring appointment series.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`year`|Integer||The year value of the end date.|
|`month`|Integer||The month value of the end date. Valid range is 0-11 where 0 represents the 1st month and 11 represents the 12th month.|
|`day`|Integer||The day value of the end date.|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the end date of a recurring appointment series to November 2, 2017.

```js
var seriesTimeObject = new seriesTime()
seriesTimeObject.setEndDate(2017, 10, 2);
```

##### Errors

|Error code|Description|
|------------|-------------|
|`Invalid date format`|The date is not in an acceptable format.|

#### setEndDate(date)

Sets the end date of a recurring appointment series.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`date`|String||End date of the recurring appointment series represented in [ISO 8601](https://www.iso.org/iso-8601-date-and-time-format.html) date standard, that is, "YYYY-MM-DD".|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the end date of a recurring appointment series to November 2, 2017.

```js
var seriesTimeObject = new seriesTime()
seriesTimeObject.setEndDate("2017-11-02");
```

##### Errors

|Error code|Description|
|------------|-------------|
|`Invalid date format`|The date is not in an acceptable format.|

#### setStartTime(hours, minutes)

Sets the start time of all instances of a recurring appointment series in whichever time zone the recurrence pattern is set (the item's time zone is used by default).

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`hours`|Integer||The `hour` value of the start time. Valid range: 0-24.|
|`minutes`|Integer||The `minute` value of the start time. Valid range: 0-59.|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the start time of each instance of a recurring appointment series to 1:30 PM.

```js
var seriesTimeObject = new seriesTime()
seriesTimeObject.setStartTime(13, 30);
```

This example sets the start time of each instance of a recurring appointment series to 11:30 AM.

```js
var seriesTimeObject = new seriesTime()
seriesTimeObject.setStartTime(11, 30);
```

##### Errors

|Error code|Description|
|------------|-------------|
|`Invalid time format`|The time is not in an acceptable format.|

#### setStartTime(time)

Sets the start time of all instances of a recurring appointment series in whichever time zone the recurrence pattern is set (the item's time zone is used by default).

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`time`|String||Start time of all instances represented by standard datetime string format, that is, "THH:mm:ss:mmm".|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the start time of all instances of a recurring appointment series to 11:30 PM.

```js
var seriesTimeObject = new seriesTime()
seriesTimeObject.setStartTime("T23:30:00");
```

##### Errors

|Error code|Description|
|------------|-------------|
|`Invalid time format`|The time is not in an acceptable format.|

#### setStartDate(year, month, day)

Sets the start date of a recurring appointment series.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`year`|Integer||The year value of the start date.|
|`month`|Integer||The month value of the start date. Valid range is 0-11 where 0 represents the 1st month and 11 represents the 12th month.|
|`day`|Integer||The day value of the start date.|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the start date of a recurring appointment series to November 2, 2017.

```js
var seriesTimeObject = new seriesTime()
seriesTimeObject.setStartDate(2017, 10, 2);
```

##### Errors

|Error code|Description|
|------------|-------------|
|`Invalid date format`|The date is not in an acceptable format.|

#### setStartDate(date)

Sets the start date of a recurring appointment series.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`date`|String||Start date of the recurring appointment series in [ISO 8601](https://www.iso.org/iso-8601-date-and-time-format.html) date standard, that is, "YYYY-MM-DD".|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the start date of a recurring appointment series to November 2, 2017.

```js
var seriesTimeObject = new seriesTime()
seriesTimeObject.setStartDate("2017-11-02");
```

##### Errors

|Error code|Description|
|------------|-------------|
|`Invalid date format`|The date is not in an acceptable format.|
