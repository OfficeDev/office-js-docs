# WorksheetDataChangedEvent Object (JavaScript API for Excel)

Provides information about the worksheet that raised the DataChanged event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|changeType|string|Gets the change type that represents how the DataChanged event is triggered. Possible values are: Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|[ApiSet.InProgressFeatures.Event](../requirement-sets/excel-api-requirement-sets.md)|
|context|object|Gets the context object that facilitates requests to the Excel application.|[ApiSet.InProgressFeatures.Event](../requirement-sets/excel-api-requirement-sets.md)|
|source|string|Gets the source of the event. Possible values are: Local, Remote.|[ApiSet.InProgressFeatures.Event](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged.|[ApiSet.InProgressFeatures.Event](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|range|[Range](range.md)|Gets the range that represents the changed area of a specific worksheet.|[ApiSet.InProgressFeatures.Event](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

