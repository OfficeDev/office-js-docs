## Introduction to ExcelEvent Features

This page describes the Excel Event APIs that are available as part of beta version. 

_**Note**: The features listed are still in the design and review phase and are not yet available as part of the product. The final design is subject to change. When the feature is made available, the final specification will be published as part of the master repository._

_Details of the specific APIs are listed below. Please let us know what you think!_

|Object| What is new| Description |Requirement set|
|:----|:----|:----|:----|
|[WorksheetChangedEventArgs](reference/excel/worksheetchangedeventargs.md)|_Property_ > address|Gets the range address that represents the changed area of a specific worksheet.|Beta|
|[WorksheetChangedEventArgs](reference/excel/worksheetchangedeventargs.md)|_Property_ > changeType|Gets the change type that represents how the DataChanged event is triggered. Possible values are: Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|Beta|
|[WorksheetChangedEventArgs](reference/excel/worksheetchangedeventargs.md)|_Property_ > source|Gets the source of the event. Possible values are: Local, Remote.|Beta|
|[WorksheetChangedEventArgs](reference/excel/worksheetchangedeventargs.md)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|Beta|
|[WorksheetChangedEventArgs](reference/excel/worksheetchangedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the data changed.|Beta|
|[WorksheetChangedEventArgs](reference/excel/worksheetchangedeventargs.md)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|Beta|
|[TableChangedEventArgs](reference/excel/tablechangedeventargs.md)|_Property_ > type |Gets the type of the event. Possible values are: WorksheetChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableChanged, TableSelectionChanged, WorksheetDeleted.|Beta|
|[TableChangedEventArgs](reference/excel/tablechangedeventargs.md)|_Property_ > changeType |Gets the change type that represents how the DataChanged event is triggered. Possible values are: Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|Beta|
|[TableChangedEventArgs](reference/excel/tablechangedeventargs.md)|_Property_ > source |Gets the source of the event. Possible values are: Local, Remote.|Beta|
|[TableChangedEventArgs](reference/excel/tablechangedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the data changed.|Beta|
|[TableChangedEventArgs](reference/excel/tablechangedeventargs.md)|_Property_ > tableId |Gets the id of the table in which the data changed.|Beta|
|[TableChangedEventArgs](reference/excel/tablechangedeventargs.md)|_Property_ > address |Gets the address that represents the changed area of a table on a specific worksheet.|Beta|
|[WorksheetSelectionChangedEventArgs](reference/excel/worksheetselectionchangedeventargs.md)|_Property_ > address|Gets the range address that represents the selected area of a specific worksheet.|Beta|
|[WorksheetSelectionChangedEventArgs](reference/excel/worksheetselectionchangedeventargs.md)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|Beta|
|[WorksheetSelectionChangedEventArgs](reference/excel/worksheetselectionchangedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the selection changed.|Beta|
|[TableSelectionChangedEventArgs](reference/excel/TableSelectionChangedEventArgs.md)|_Property_ > type |Gets the type of the event. Possible values are: WorksheetChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableChanged, TableSelectionChanged, WorksheetDeleted.|Beta|
|[TableSelectionChangedEventArgs](reference/excel/TableSelectionChangedEventArgs.md)|_Property_ > IsInsideTable |Indicates if the selection is inside a table, address will be useless if IsInsideTable is false.|Beta|
|[TableSelectionChangedEventArgs](reference/excel/TableSelectionChangedEventArgs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the selection changed.|Beta|
|[TableSelectionChangedEventArgs](reference/excel/TableSelectionChangedEventArgs.md)|_Property_ > tableId |Gets the id of the table in which the selection changed.|Beta|
|[TableSelectionChangedEventArgs](reference/excel/TableSelectionChangedEventArgs.md)|_Property_ > address |Gets the address that represents the changed area of a table on a specific worksheet.|Beta|
[WorksheetAddedEventArgs](reference/excel/WorksheetAddedEventArgs.md)|_Property_ > source|Gets the source of the event. Possible values are: Local, Remote.|Beta|
|[WorksheetAddedEventArgs](reference/excel/WorksheetAddedEventArgs.md)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|Beta|
|[worksheetAddedEvent](reference/excel/WorksheetAddedEventArgs.md)|_Property_ > worksheetId|Gets the id of the worksheet that is added to the workbook.|Beta|
|[WorksheetActivatedEventArgs](reference/excel/worksheetactivatedeventargs.md)|_Property_ > WorksheetId|Gets the id of the worksheet that is activated.|Beta|
|[WorksheetActivatedEventArgs](reference/excel/worksheetactivatedeventargs.md)|_Property_ > Type|Gets the type of the event. Possible values include: WorksheetChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableChanged, TableSelectionChanged.|Beta|
|[WorksheetDeactivatedEventArgs](reference/excel/worksheetdeactivatedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet that is deactivated.|Beta|
|[WorksheetDeactivatedEventArgs](reference/excel/worksheetdeactivatedeventargs.md)|_Property_ > type|Gets the type of the event. Possible values include: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged.|Beta|



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

function onWorksheetDataChanged(eventargs) {
    return Excel.run(eventargs, function (context) {
        var worksheet = eventargs.worksheet;
        var range = eventargs.range;
        var source = eventargs.source;
        var changeType = eventargs.changeType;

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
