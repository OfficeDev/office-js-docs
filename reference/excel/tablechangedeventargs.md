# TableChangedEventArgs Object (JavaScript API for Excel)

Provides information about the table that raised the DataChanged event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|type|string|Gets the type of the event. Possible values are: WorksheetChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableChanged, TableSelectionChanged, WorksheetDeleted.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|changeType|string|Gets the change type that represents how the DataChanged event is triggered. Possible values are: Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|source|string|Gets the source of the event. Possible values are: Local, Remote.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|worksheetId|string|Gets the id of the worksheet in which the data changed.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|tableId|string|Gets the id of the table in which the data changed.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|address|string|Gets the address that represents the changed area of a table on a specific worksheet.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|



_See property access [examples.](#property-access-examples)_


