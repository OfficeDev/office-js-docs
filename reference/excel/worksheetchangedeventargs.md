# WorksheetChangedEventArgs Object (JavaScript API for Excel)

Provides information about the worksheet that raised the DataChanged event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|Type|string|Gets the type of the event. Possible values include: WorksheetChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableChanged, TableSelectionChanged. |[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|ChangeType|string|Gets the change type that represents how the DataChanged event is triggered. Possible values include: RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted, Others.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|Source|string|Gets the source of the event. Possible values are: Local, Remote.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|WorksheetId|string|Gets the id of the worksheet in which the data changed.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|Address|string|Gets the range address that represents the changed area of a specific worksheet.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
