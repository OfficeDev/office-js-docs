# TableDataChangedEvent Object (JavaScript API for Excel)

Provides information about the table that raised the DataChanged event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|context|object|Gets the context object that facilitates requests to the Excel application.|[1.7E](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|changeType|string|Gets the change type that represents how the DataChanged event is triggered.|[1.7E](../requirement-sets/excel-api-requirement-sets.md)|
|range|[Range](range.md)|Gets the range that represents the changed area of a table on a specific worksheet.|[1.7E](../requirement-sets/excel-api-requirement-sets.md)|
|source|string|Gets the source of the event.|[1.7E](../requirement-sets/excel-api-requirement-sets.md)|
|table|[Table](table.md)|Gets the table in which the data changed.|[1.7E](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Gets the type of the event.|[1.7E](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

