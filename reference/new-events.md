# Event API


## New Events in Beta

| Object | Event | Description | Event Argument |
|:----|:----|:----|:---|
| Chart | onActivated | Occurs when the chart has become activated. | [ChartActivatedEventArgs](/reference/excel/chartactivatedeventargs.md) |
| Chart | onDeactivated | Occurs when the chart has become deactivated. | [ChartDeactivatedEventArgs](/reference/excel/chartdeactivatedeventargs.md) |
| ChartCollection | onActivated | Occurs when any chart has become activated. | [ChartActivatedEventArgs](/reference/excel/chartactivatedeventargs.md) |
| ChartCollection | onDeactivated | Occurs when any chart has become deactivated | [ChartDeactivatedEventArgs](/reference/excel/chartdeactivatedeventargs.md) |
| ChartCollection | onAdded | Occurs when a chart has been added. | [ChartAddedEventArgs](/reference/excel/chartaddedeventargs.md) |
| ChartCollection | onDeleted | Occurs when a worksheet has been deleted. | [ChartDeletedEventArgs](/reference/excel/chartdeletedeventargs.md) |
| Worksheet | onCalculated | Occurs when the workbook has finished calculation. | [WorsheetCalculatedEventArgs](/reference/excel/worksheetcalculatedeventargs.md) |
 WorksheetCollection | onCalculated | Occurs when all the worksheets of the workbook have finished calculation. | [WorsheetCalculatedEventArgs](/reference/excel/worksheetcalculatedeventargs.md) |

## Upcoming Events to be preivewed

| Object | Event | Description | Event Argument |
|:----|:----|:----|:---|
| WorksheetCollection | onChanged | Occurs when cells in any worksheet are changed. Chart sheet isn’t included. | [WorksheetChangedEventArgs](/reference/excel/worksheetchangedeventargs.md)|
| WorksheetCollection | onSelectionChanged | Occurs when the selection changes on any worksheet. Chart sheet isn’t included. | [WorsheetCalculatedEventArgs](/reference/excel/worksheetselectionchangedeventargs.md)|
| TableCollection | onAdded | Occurs when a table is added. | TableCollectionAddedEventArgs|
| TableCollection | onDeleted | Occurs when a table is deleted. | TableCollectionDeletedEventArgs|

## Change log

| Object | Content| Change Type | Description | 
|:----|:----|:----|:---|
| [WorksheetChangedEventArgs](/reference/excel/worksheetchangedeventargs.md) | getRange() getRangeOrNullObject()| New Method | Eventargs.address may become out of date.(e.g. insert a row before where the change happened) A temparory dynamic range object is created in the background which will be kept up to 10 sec. This mehod returns a persisted copy of this dynamic range object.|
| [Runtime](/reference/excel/runtime.md) | enableEvents |  New Property | Turn JavaScript events on and off for the current taskpane or content add-in. Do sync() before calling other APIs to make it effect. |
