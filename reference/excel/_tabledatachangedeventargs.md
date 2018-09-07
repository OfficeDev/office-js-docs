# _TableDataChangedEventArgs Object (JavaScript API for Excel)

Provides information about the worksheet that raised the Changed event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|address|string|Gets the range address that represents the changed area of a specific worksheet.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|referenceId|string|Gets the range address that represents the changed area of a specific worksheet.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|tableId|string|Gets the range address that represents the changed area of a specific worksheet.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|worksheetId|string|Gets the range address that represents the changed area of a specific worksheet.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|changeType|[DataChangeType](datachangetype.md)|Gets the range address that represents the changed area of a specific worksheet.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|source|[EventSource](eventsource.md)|Gets the range address that represents the changed area of a specific worksheet.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

