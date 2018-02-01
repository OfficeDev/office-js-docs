# WorksheetAddedEventArgs Object (JavaScript API for Excel)

Provides information about the worksheet that raised the Added event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|Type|string|Gets the type of the event. Possible values include: WorksheetChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableChanged, TableSelectionChanged.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|Source|string|Get the source of the event. Possible values include: Local, Remote.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|WorksheetId|string|Gets the id of the worksheet that is added to the workbook.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|

