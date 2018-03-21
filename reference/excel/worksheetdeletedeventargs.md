# WorksheetDeletedEventArgs Object (JavaScript API for Excel)

Provides information about the worksheet that raised the Deleted event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|source|string|Gets the source of the event. Possible values are: Local, Remote.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|worksheetId|string|Gets the id of the worksheet that is deleted from the workbook.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods
None

