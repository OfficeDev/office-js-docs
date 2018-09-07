# _WorksheetDataChangedEventArgs Object (JavaScript API for Excel)

Provides information about arguments of WorksheetDataChangedEvent returned from host.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|address|string|Gets the Workbook object that raised the Message event|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|referenceId|string|Gets the Workbook object that raised the Message event|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|worksheetId|string|Gets the Workbook object that raised the Message event|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|changeType|[DataChangeType](datachangetype.md)|Gets the Workbook object that raised the Message event|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|source|[EventSource](eventsource.md)|Gets the Workbook object that raised the Message event|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

