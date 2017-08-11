## Introduction to Excel 1.8 Event Features

This page describes the Excel Event APIs that are planned for the next release (requirement set 1.8). 

_**Note**: The features listed are still in the design and review phase and are not yet available as part of the product. The final design is subject to change. When the feature is made available, the final specification will be published as part of the master repository._

_Details of the specific APIs are listed below. Please let us know what you think!_

|Object| What is new| Description |Requirement set|
|:----|:----|:----|:----|
|[Table](reference/excel/table.md)|_Event_ > [onDataChanged](reference/excel/table.md#getfirst)|Occurs when data changed on a specific table.|1.8|
|[Table](reference/excel/table.md)|_Event_ > [onSelectionChanged](reference/excel/table.md#getfirst)|Occurs when the selection changed on a specific table.|1.8|
|[TableCollection](reference/excel/tablecollection.md)|_Event_ > [onDataChanged](reference/excel/tablecollection.md#getfirst)|Occurs when data changed on any table in a workbook, or a worksheet.|1.8|
|[Worksheet](reference/excel/worksheet.md)|_Event_ > [onActivated](reference/excel/worksheet.md#getfirst)|Occurs when the worksheet is activated.|1.8|
|[Worksheet](reference/excel/worksheet.md)|_Event_ > [onDataChanged](reference/excel/worksheet.md#getfirst)|Occurs when data changed on a specific worksheet.|1.8|
|[Worksheet](reference/excel/worksheet.md)|_Event_ > [onDeactivated](reference/excel/worksheet.md#getfirst)|Occurs when the worksheet is deactivated.|1.8|
|[Worksheet](reference/excel/worksheet.md)|_Event_ > [onSelectionChanged](reference/excel/worksheet.md#getfirst)|Occurs when the selection changed on a specific worksheet.|1.8|
|[WorksheetCollection](reference/excel/worksheetcollection.md)|_Event_ > [onActived](reference/excel/worksheetcollection.md#getfirst)|Occurs when any worksheet in the workbook is activated.|1.8|
|[WorksheetCollection](reference/excel/worksheetcollection.md)|_Event_ > [onAdded](reference/excel/worksheetcollection.md#getfirst)|Occurs when a new worksheet is added to the workbook.|1.8|
|[WorksheetCollection](reference/excel/worksheetcollection.md)|_Event_ > [onDeactivated](reference/excel/worksheetcollection.md#getlast)|Occurs when any worksheet in the workbook is deactivated.|1.8|

### Examples

```js
/** Worksheet DataChanged event */
Excel.run(function (ctx) {
    ctx.workbook.worksheets.getItem(“Sheet1”).onDataChanged.add(onWorksheetDataChanged);
    return ctx.sync().then(function () {
        console.log("add event succeed.");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});

function onWorksheetDataChanged(event) {
    return Excel.run(event, function (context) {
        var worksheet = event.worksheet;
        var range = event.range;
        var source = event.source;
        var changeType = event.changeType;

        worksheet.load("name");
        range.load("address, cellCount");

        return context.sync().then(function () {
            // Get the changed range
            console.log("Data Changed in worksheet " + worksheet.name + ".Changed area address is " + range.address + "cell count in the range is" + range.cellCount);
            //Get the event source, data change comes from local or remote.
            console.log("Data Change comes from" + source);
          
            //get the new inserted Row.
            else if (changeType == "RowInsert") {
                console.log("Row inserted on row" + range.address);
            }
            //else if ... //For RowDelete, ColumnInsert, ColumnDelete, same as above.

            // Other changes happened
            else {
                console.log("Other changes happened");
            }
        });
    });
}
```
