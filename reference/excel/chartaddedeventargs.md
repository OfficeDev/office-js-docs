# ChartAddedEventArgs Object (JavaScript API for Excel)

Provides information about the chart that raised the Added event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|chartId|string|Gets the id of the chart that is added to the worksheet.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|source|string|Gets the source of the event. Possible values are: Local, Remote.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|worksheetId|string|Gets the id of the worksheet in which the chart is added.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods
None

