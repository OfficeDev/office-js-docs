# WorksheetAddedEvent Object (JavaScript API for Excel)

Provides information about the worksheet that raised the Added event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|context|object|Gets the context object that facilitates requests to the Excel application.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|source|string|Gets the source of the event. Possible values are: Local, Remote.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|worksheet|[Worksheet](worksheet.md)|Gets the worksheet that is added to the workbook.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

