# WorksheetDeletedEvent Object (JavaScript API for Excel)

Provides information about the worksheet that raised the Deleted event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|Type|string|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|Source|string|Gets the source of the event. Possible values are: Local, Remote.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|WorksheetId|string|Gets the id of the worksheet that is added to the workbook.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_


