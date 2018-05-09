

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
// Perhaps set first?

getDuration();

// example result: 120
```

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### getEndDate() → {String}

Gets the end date of a recurrence pattern in the following ISO format:
"YYYY-MM-DD"

##### Examples

This example gets the end date of a recurrence.

```js
// Perhaps set first?

getEndDate();

// example result: "2017-11-2"
```

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### getEndTime() → {String}

Gets the start time of the meeting/appointment of a recurrence pattern in whichever time zone that the user/add-in set the recurrence pattern in in the following ISO format:
"THH:MM:SS:mmm"

##### Examples

This example gets the end time of a recurrence.

```js
// Perhaps set first?

getEndTime();

// example result: "EXAMPLE"
```

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### getStartDate() → {String}

Gets the start date of a recurrence pattern in the following ISO format:
"YYYY-MM-DD"

##### Examples

This example gets the start date of a recurrence.

```js
// Perhaps set first?

getStartDate();

// example result: "2017-11-2"
```

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|Applicable Outlook mode|Compose or read|

#### getStartTime() → {String}

Gets the start time of the meeting/appointment of a recurrence pattern in whichever time zone that the user/add-in set the recurrence pattern in in the following ISO format:
"THH:MM:SS:mmm"

##### Examples

This example gets the start time of a recurrence.

```js
// Perhaps set first?

getStartTime();

// example result: "EXAMPLE"
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
|`minutes`|Integer|||

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the duration of a recurrence to 2 hours.

```js
setDuration(120);
```

#### setEndDate(int year, int month, int day)

Sets the end date of a recurrence pattern.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`year`|Integer|||
|`month`|Integer||First month is 0.|
|`day`|Integer|||

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the end date of a recurrence to November 2, 2017.

```js
setEndDate(2017, 10, 2);
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
|`date`|String||End date of the recurrence represented in ISO-8601 standard, that is, "YYYY-MM-DD".|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the end date of a recurrence to November 2, 2017.

```js
setEndDate("2017-11-02");
```

##### Errors

|Error code|Description|
|------------|-------------|
|`Invalid date format`||

#### setStartTime(int hours, int minutes)

Sets the start time of the meeting/appointment of a recurrence pattern in whichever time zone the recurrence pattern is set in. Default value is the item's time zone.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`hours`|Integer||0-24.|
|`minutes`|Integer||0-59.|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the start time of a recurrence to 1:30 pm.

```js
setStartTime(13, 30);
```

This example sets the start time of a recurrence to 11:30 am.

```js
setStartTime(11, 30);
```

##### Errors

|Error code|Description|
|------------|-------------|
|`Invalid time format`||

#### setStartTime(String time)

Sets the start time of the meeting/appointment of a recurrence pattern in whichever time zone the recurrence pattern is set in. Default value is the item's time zone.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`time`|String||Start time of the recurrence represented in standard datetime string format, that is, "THH:MM:SS:mmm".|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the start time of a recurrence to 11:30 pm.

```js
setStartTime("ADD EXAMPLE");
```

##### Errors

|Error code|Description|
|------------|-------------|
|`Invalid time format`||

#### setStartDate(int year, int month, int day)

Sets the start date of a recurrence pattern.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`year`|Integer|||
|`month`|Integer||First month is 0.|
|`day`|Integer|||

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the start date of a recurrence to November 2, 2017.

```js
setStartDate(2017, 10, 2);
```

##### Errors

|Error code|Description|
|------------|-------------|
|`Invalid date format`||

#### setStartDate(String date)

Sets the start date of a recurrence pattern.

##### Parameters:

|Name|Type|Attributes|Description|
|---|---|---|---|
|`date`|String||Start date of the recurrence in ISO-8601 standard, that is, "YYYY-MM-DD".|

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|Applicable Outlook mode|Compose|

##### Examples

This example sets the start date of a recurrence to November 2, 2017.

```js
setStartDate("2017-11-02");
```

##### Errors

|Error code|Description|
|------------|-------------|
|`Invalid date format`||