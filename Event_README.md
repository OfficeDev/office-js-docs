## Introduction to Excel 1.8 Event Features

This page describes the Excel Event APIs that are planned for the next release (requirement set 1.8). 

_**Note**: The features listed are still in the design and review phase and are not yet available as part of the product. The final design is subject to change. When the feature is made available, the final specification will be published as part of the master repository._

_Details of the specific APIs are listed below. Please let us know what you think!_


|Object| What is new| Description |Requirement set|
|:----|:----|:----|:----|
|[Table](reference/excel/table.md)|_Event_ > [onDataChanged](reference/excel/table.md#ondatachanged)|Occurs when data changed on a specific table.|1.8|
|[Table](reference/excel/table.md)|_Event_ > [onSelectionChanged](reference/excel/table.md#onselectionchanged)|Occurs when the selection changed on a specific table.|1.8|
|[TableCollection](reference/excel/tablecollection.md)|_Event_ > [onDataChanged](reference/excel/tablecollection.md#ondatachanged)|Occurs when data changed on any table in a workbook, or a worksheet.|1.8|
|[Worksheet](reference/excel/worksheet.md)|_Event_ > [onActivated](reference/excel/worksheet.md#onactivated)|Occurs when the worksheet is activated.|1.8|
|[Worksheet](reference/excel/worksheet.md)|_Event_ > [onDataChanged](reference/excel/worksheet.md#ondatachanged)|Occurs when data changed on a specific worksheet.|1.8|
|[Worksheet](reference/excel/worksheet.md)|_Event_ > [onDeactivated](reference/excel/worksheet.md#ondeactivated)|Occurs when the worksheet is deactivated.|1.8|
|[Worksheet](reference/excel/worksheet.md)|_Event_ > [onSelectionChanged](reference/excel/worksheet.md#onselectionchanged)|Occurs when the selection changed on a specific worksheet.|1.8|
|[WorksheetCollection](reference/excel/worksheetcollection.md)|_Event_ > [onActivated](reference/excel/worksheetcollection.md#onactivated)|Occurs when any worksheet in the workbook is activated.|1.8|
|[WorksheetCollection](reference/excel/worksheetcollection.md)|_Event_ > [onAdded](reference/excel/worksheetcollection.md#onadded)|Occurs when a new worksheet is added to the workbook.|1.8|
|[WorksheetCollection](reference/excel/worksheetcollection.md)|_Event_ > [onDeactivated](reference/excel/worksheetcollection.md#ondeactivated)|Occurs when any worksheet in the workbook is deactivated.|1.8|
