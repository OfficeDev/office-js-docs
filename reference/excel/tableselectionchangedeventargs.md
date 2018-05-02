# TableSelectionChangedEventArgs Object (JavaScript API for Excel)

Provides information about the table that raised the SelectionChanged event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|address|string|Gets the range address that represents the selected area of the table on a specific worksheet.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|isInsideTable|bool|Indicates if the selection is inside a table, address will be useless if IsInsideTable is false.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|tableId|string|Gets the id of the table in which the selection changed.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Gets the type of the event. Possible values are: Value is `TableSelectionChanged`, read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|worksheetId|string|Gets the id of the worksheet in which the selection changed.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_


## Methods
None

