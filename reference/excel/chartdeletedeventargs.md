# ChartDeletedEventArgs Object (JavaScript API for Excel)

Provides information about the chart that raised the Deleted event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|chartId|string|Gets the id of the chart that is deleted from the worksheet.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|source|string|Gets the source of the event. Possible values are: Local, Remote.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Gets the type of the event. Value is `ChartDeleted`, read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|worksheetId|string|Gets the id of the worksheet in which the chart is deleted.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_


## Methods
None

