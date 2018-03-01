# WorksheetDeletedEventArgs Object (JavaScript API for Excel)

Provides information about the worksheet that raised the Deleted event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|Type|string|Gets the type of the event. Possible values are: WorksheetChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableChanged, TableSelectionChanged.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|Source|string|Gets the source of the event. Possible values are: Local, Remote.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|WorksheetId|string|Gets the id of the worksheet that is added to the workbook.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Examples
```js
Excel.run(function (ctx) {  
    var worksheets = ctx.workbook.worksheets; 
    worksheets.onDeleted.add(onWorksheetDeleted); 
    return ctx.sync().then(function() { 
                    console.log("add event success."); 
    }); 
}).catch(function(error) { 
        console.log("Error: " + error); 
        if (error instanceof OfficeExtension.Error) { 
            console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
        } 
}); 
 
function onWorksheetDeleted(event)
{ 
        console.log("delete worksheet id is : " + event.worksheetId);
}
```
