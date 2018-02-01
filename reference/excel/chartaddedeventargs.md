# ChartAddedEventArgs Object (JavaScript API for Excel)

Provides information about the worksheet that raised the Added event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|Type|string|Get the type of the event. See Excel.Event.Type for details.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|Source|string|Get the source of the event. Possible values include: Local, Remote.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|WorksheetId|string|Gets the id of the chart that is added to the worksheet.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|ChartId|string|Gets the id of the chart that is added to the worksheet.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

