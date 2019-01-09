# WorksheetChangedEventArgs Object (JavaScript API for Excel)

Provides information about the worksheet that raised the Changed event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|address|string|Gets the range address that represents the changed area of a specific worksheet.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Gets the type of the event.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|worksheetId|string|Gets the id of the worksheet in which the data changed.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|changeType|string|Gets the change type that represents how the Changed event is triggered.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|source|string|Gets the source of the event.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

