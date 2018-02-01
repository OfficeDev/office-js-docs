# WorksheetSelectionChangedEventArgs Object (JavaScript API for Excel)

Provides information about the worksheet that raised the SelectionChanged event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|Type|string|Gets the type of the event. Possible values are: WorksheetChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableChanged, TableSelectionChanged, WorksheetDeleted.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|WorksheetId|string|Gets the id of the worksheet in which the selection changed.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|Address|string|Gets the range address that represents the selected area of a specific worksheet.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|

