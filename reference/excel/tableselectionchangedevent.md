# TableSelectionChangedEvent Object (JavaScript API for Excel)

Provides information about the table that raised the SelectionChanged event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|context|object|Gets the context object that facilitates requests to the Excel application.|[ApiSet.InProgressFeatures.Event](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged.|[ApiSet.InProgressFeatures.Event](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|range|[Range](range.md)|Gets the range that represents the selected area of the table on a specific worksheet.|[ApiSet.InProgressFeatures.Event](../requirement-sets/excel-api-requirement-sets.md)|
|table|[Table](table.md)|Gets the table in which the selection changed.|[ApiSet.InProgressFeatures.Event](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

