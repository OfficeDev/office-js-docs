## Introduction to Excel 1.8 Event Features

This page describes the Excel Event APIs that are planned for the next release (requirement set 1.8). 

_**Note**: The features listed are still in the design and review phase and are not yet available as part of the product. The final design is subject to change. When the feature is made available, the final specification will be published as part of the master repository._

_Details of the specific APIs are listed below. Please let us know what you think!_

|Object| What is new| Description |Requirement set|
|:----|:----|:----|:----|
|[registerEventPostProcessAction](reference/excel/registereventpostprocessaction.md)|_Property_ > actionType|Gets the top border Possible values are: None, RegisterEvent, UnregisterEvent, CustomFunction, ShowCard.|TBD|
|[registerEventPostProcessAction](reference/excel/registereventpostprocessaction.md)|_Property_ > message|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent.|TBD|
|[registerEventPostProcessAction](reference/excel/registereventpostprocessaction.md)|_Property_ > messageType|Gets the top border Possible values are: None, RegisterEvent, UnregisterEvent, CustomFunction, ShowCard.|TBD|
|[registerEventPostProcessAction](reference/excel/registereventpostprocessaction.md)|_Property_ > targetId|Gets the top border Possible values are: None, RegisterEvent, UnregisterEvent, CustomFunction, ShowCard.|TBD|
|[registerEventPostProcessAction](reference/excel/registereventpostprocessaction.md)|_Relationship_ > controlId|Gets the top border Possible values are: None, RegisterEvent, UnregisterEvent, CustomFunction, ShowCard.|TBD|
|[tableDataChangedEvent](reference/excel/tabledatachangedevent.md)|_Property_ > address|Gets the address that represents the changed area of a table on a specific worksheet.|TBD|
|[tableDataChangedEvent](reference/excel/tabledatachangedevent.md)|_Property_ > changeType|Gets the change type that represents how the DataChanged event is triggered. Possible values are: Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|TBD|
|[tableDataChangedEvent](reference/excel/tabledatachangedevent.md)|_Property_ > source|Gets the source of the event. Possible values are: Local, Remote.|TBD|
|[tableDataChangedEvent](reference/excel/tabledatachangedevent.md)|_Property_ > tableId|Gets the id of the table in which the data changed.|TBD|
|[tableDataChangedEvent](reference/excel/tabledatachangedevent.md)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|TBD|
|[tableDataChangedEvent](reference/excel/tabledatachangedevent.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the data changed.|TBD|
|[tableSelectionChangedEvent](reference/excel/tableselectionchangedevent.md)|_Property_ > address|Gets the range address that represents the selected area of the table on a specific worksheet.|TBD|
|[tableSelectionChangedEvent](reference/excel/tableselectionchangedevent.md)|_Property_ > tableId|Gets the id of the table in which the selection changed.|TBD|
|[tableSelectionChangedEvent](reference/excel/tableselectionchangedevent.md)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|TBD|
|[tableSelectionChangedEvent](reference/excel/tableselectionchangedevent.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the selection changed.|TBD|
|[worksheetActivatedEvent](reference/excel/worksheetactivatedevent.md)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|TBD|
|[worksheetActivatedEvent](reference/excel/worksheetactivatedevent.md)|_Property_ > worksheetId|Gets the id of the worksheet that is activated.|TBD|
|[worksheetAddedEvent](reference/excel/worksheetaddedevent.md)|_Property_ > source|Gets the source of the event. Possible values are: Local, Remote.|TBD|
|[worksheetAddedEvent](reference/excel/worksheetaddedevent.md)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|TBD|
|[worksheetAddedEvent](reference/excel/worksheetaddedevent.md)|_Property_ > worksheetId|Gets the id of the worksheet that is added to the workbook.|TBD|
|[worksheetDataChangedEvent](reference/excel/worksheetdatachangedevent.md)|_Property_ > address|Gets the range address that represents the changed area of a specific worksheet.|TBD|
|[worksheetDataChangedEvent](reference/excel/worksheetdatachangedevent.md)|_Property_ > changeType|Gets the change type that represents how the DataChanged event is triggered. Possible values are: Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|TBD|
|[worksheetDataChangedEvent](reference/excel/worksheetdatachangedevent.md)|_Property_ > source|Gets the source of the event. Possible values are: Local, Remote.|TBD|
|[worksheetDataChangedEvent](reference/excel/worksheetdatachangedevent.md)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|TBD|
|[worksheetDataChangedEvent](reference/excel/worksheetdatachangedevent.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the data changed.|TBD|
|[worksheetDeactivatedEvent](reference/excel/worksheetdeactivatedevent.md)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|TBD|
|[worksheetDeactivatedEvent](reference/excel/worksheetdeactivatedevent.md)|_Property_ > worksheetId|Gets the id of the worksheet that is deactivated.|TBD|
|[worksheetDeletedEvent](reference/excel/worksheetdeletedevent.md)|_Property_ > source|Gets the source of the event. Possible values are: Local, Remote.|TBD|
|[worksheetDeletedEvent](reference/excel/worksheetdeletedevent.md)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|TBD|
|[worksheetDeletedEvent](reference/excel/worksheetdeletedevent.md)|_Property_ > worksheetId|Gets the id of the worksheet that is added to the workbook.|TBD|
|[worksheetSelectionChangedEvent](reference/excel/worksheetselectionchangedevent.md)|_Property_ > address|Gets the range address that represents the selected area of a specific worksheet.|TBD|
|[worksheetSelectionChangedEvent](reference/excel/worksheetselectionchangedevent.md)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|TBD|
|[worksheetSelectionChangedEvent](reference/excel/worksheetselectionchangedevent.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the selection changed.|TBD|

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
