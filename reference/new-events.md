# Event API


## New Events in Beta

| Object | Event | Description | Event Argument |
|:----|:----|:----|:---|
| Chart | onActivated | Occurs when the chart has become activated. | [ChartActivatedEventArgs](/reference/excel/chartactivatedeventargs.md) |
| Chart | onDeactivated | Occurs when the chart has become deactivated. | [ChartDeactivatedEventArgs](/reference/excel/chartdeactivatedeventargs.md) |
| ChartCollection | onActivated | Occurs when any chart has become activated. | [ChartActivatedEventArgs](/reference/excel/chartactivatedeventargs.md) |
| ChartCollection | onDeactivated | Occurs when any chart has become deactivated | [ChartDeactivatedEventArgs](/reference/excel/chartdeactivatedeventargs.md) |
| ChartCollection | onAdded | Occurs when a chart has been added. | [ChartAddedEventArgs](/reference/excel/chartaddedeventargs.md) |
| ChartCollection | onDeleted | Occurs when a worksheet has been deleted from the workbook. | [ChartDeletedEvent](/reference/excel/chartdeletedeventargs.md) |
| Worksheet | onCalculated | Occurs when the workbook has finished calculation. | [WorsheetCalculatedEventArgs](/reference/excel/worksheetcalculatedeventargs.md) |
 WorksheetCollection | onCalculated | Occurs when all the worksheets of the workbook have finished calculation. | [WorsheetCalculatedEventArgs](/reference/excel/worksheetcalculatedeventargs.md) |

## Upcoming Events to be preivewed

| Object | Event | Description | Event Argument |
|:----|:----|:----|:---|
| WorksheetCollection | onChanged | Occurs when cells in any worksheet are changed. Chart sheet isn’t included. | [WorksheetChangedEventAargs](/reference/excel/worksheetchangedeventargs.md)|
| WorksheetCollection | onSelectionChanged | Occurs when the selection changes on any worksheet. Chart sheet isn’t included. | [WorsheetCalculatedEventArgs](/reference/excel/worksheetselectionchangedeventargs.md)|
| TableCollection | onAdded | Occurs when a table is added on worksheet. | TableCollectiAddedEventArgs|
| TableCollection | onDeleted | Occurs when a table is deleted on worksheet. | TableCollectiDeletedEventArgs|

## Change log

| Object | Content| Type | Description | 
|:----|:----|:----|:---|
| [worksheetchangedeventargs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetchangedeventargs.md) | getRange() | New Method | Return the range object that is associated with the address when the event occurs.|
| context.runtime | enableEvents() |  New Method | Turn JavaScript events on and off for the current taskpane or content add-in. |
