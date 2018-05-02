# WorksheetChangedEventArgs Object (JavaScript API for Excel)

Provides information about the worksheet that raised the Changed event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|address|string|Gets the range address that represents the changed area of a specific worksheet.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|changeType|string|Gets the change type that represents how the Changed event is triggered. Possible values are: Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|source|string|Gets the source of the event. Possible values are: Local, Remote.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|worksheetId|string|Gets the id of the worksheet in which the data changed.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_


## Methods
| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|getRange() gerRangeOrNullObject()|range|Eventargs.address may become out of date.(e.g. insert a row before where the change happened) A temparory dynamic range object is created in the background which will be kept up to 10 sec. This mehod returns a persisted copy of this dynamic range object.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Examples


### getRange() Examples
```js
function onWorksheetChanged(eventArgs) {
    return Excel.run(async function (context) {
        var range = eventArgs.getRange(context);
        range.load("address");
        await context.sync();
        
        console.log("Address:" + range.address)
    });
}
```

