# PivotTableExt Object (JavaScript API for Excel)

Represents a PivotTable report on a worksheet.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|allowMultipleFilters|bool|Sets or retrieves a value that indicates whether a PivotField can have multiple filters applied to it at the same time.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|alternativeText|string|Returns or sets the descriptive (alternative) text string for the specified PivotTable.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|columnGrandTotals|bool|True if the PivotTable report shows grand totals for columns.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|compactLayoutColumnHeader|string|Specifies the caption that is displayed in the column header of a PivotTable when in compact row layout form.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|compactLayoutRowHeader|string|Specifies the caption that is displayed in the row header of a PivotTable when in compact row layout form.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|compactRowIndent|int|Returns or sets the indent increment for PivotItems when compact row layout form is turned on.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|displayContextTooltips|bool|Controls whether or not tooltips are displayed for PivotTable cells.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|displayEmptyColumn|bool|Returns True when the non-empty MDX keyword is included in the query to the OLAP provider for the value axis. The OLAP provider will not return empty columns in the result set. Returns False when the non-empty keyword is omitted.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|displayEmptyRow|bool|Returns True when the non-empty MDX keyword is included in the query to the OLAP provider for the category axis. The OLAP provider will not return empty rows in the result set. Returns False when the non-empty keyword is omitted.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|displayErrorString|bool|True if the PivotTable report displays a custom error string in cells that contain errors. The default value is False.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|displayFieldCaptions|bool|Controls whether or not filter buttons and PivotField captions for rows and columns are displayed in the grid.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|displayNullString|bool|True if the PivotTable report displays a custom string in cells that contain null values. The default value is True.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|enableDrilldown|bool|True if drilldown is enabled. The default value is True.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|enableFieldDialog|bool|True if the PivotTable Field dialog box is available when the user double-clicks the PivotTable field. The default value is True.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|enableFieldList|bool|False to disable the ability to display the field list for the PivotTable. If the field list was already being displayed it disappears. The default value is True.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|enableWizard|bool|True if the PivotTable Wizard is available. The default value is True.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|errorString|string|Returns or sets a String value that represents the string displayed in cells that contain errors when the DisplayErrorString property is True.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|fieldListSortAscending|bool|Controls the sort order of fields in the PivotTable Field List. When this property is set toTrue, the fields are sorted in ascending order. When it is set to False, the fields are sorted in data source order.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|grandTotalName|string|Returns or sets the text string label that is displayed in the grand total column or row heading in the specified PivotTable report. The default value is the string `Grand Total`.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|hasAutoFormat|bool|True if the PivotTable report is automatically formatted when it╬ô├ç├ûs refreshed or when fields are moved.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|hidden|bool|Checks whether the PivotTable exists at the worksheet level. Boolean. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|nullString|string|Returns or sets the string displayed in cells that contain null values when theDisplayNullString property is True. The default value is an empty string.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|preserveFormatting|bool|True if formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.For query tables, this property is True if any formatting common to the first five rows of data are applied to new rows of data in the query table. Unused cells aren╬ô├ç├ût formatted. The property isFalse if the last AutoFormat applied to the query table is applied to new rows of data. The default value is True.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|printDrillIndicators|bool|Specifies whether or not drill indicators are printed with the PivotTable.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|printTitles|bool|True if the print titles for the worksheet are set based on the PivotTable report. False if the print titles for the worksheet are used. The default value is False.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|refreshName|string|Returns the name of the person who last refreshed the PivotTable report data. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|repeatItemsOnEachPrintedPage|bool|True if row, column, and item labels appear on the first row of each page when the specified PivotTable report is printed. False if labels are printed only on the first page. The default value is True.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowGrandTotals|bool|True if the PivotTable report shows grand totals for rows.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|saveData|bool|True if data for the PivotTable report is saved with the workbook. False if only the report definition is saved.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showDrillIndicators|bool|The ShowDrillIndicators property is used for toggling the display of drill indicators in the PivotTable.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showPageMultipleItemLabel|bool|When set to True (default), Multiple items will appear in the PivotTable cell on the worksheet whenever items are hidden and an aggregate of non-hidden items is shown in the PivotTable view.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showTableStyleColumnHeaders|bool|The ShowTableStyleColumnHeaders property is set to True if the coulmn headers should be displayed in the PivotTable.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showTableStyleColumnStripes|bool|The ShowTableStyleColumnStripes property displays banded columns in which even columns are formatted differently from odd columns. This makes PivotTables easier to read.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showTableStyleLastColumn|bool|Returns or sets if the last column is displayed for the specified PivotTable object.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showTableStyleRowHeaders|bool|The ShowTableStyleRowHeaders property is set to True if the row headers should be displayed in the PivotTable.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showTableStyleRowStripes|bool|The ShowTableStyleRowStripes property displays banded rows in which even rows are formatted differently from odd rows. This makes PivotTables easier to read.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showValuesRow|bool|Returns or sets whether the values row is displayed.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|smallGrid|bool|True if Microsoft Excel uses a grid that╬ô├ç├ûs two cells wide and two cells deep for a newly created PivotTable report. False if Excel uses a blank stencil outline.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|sortUsingCustomLists|bool|The SortUsingCustomLists property controls whether custom lists are used for sorting items of fields, both initially when the PivotField is initialized and the PivotItems are ordered by their captions; and later when the user applies a sort.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|subtotalHiddenPageItems|bool|True if hidden page field items in the PivotTable report are included in row and column subtotals, block totals, and grand totals. The default value is False.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|summary|string|Returns or sets the description associated with the alternative text string for the specified PivotTable.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|tag|string|Returns or sets a string saved with the PivotTable report.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|totalsAnnotation|bool|True if an asterisk (*) is displayed next to each subtotal and grand total value in the specified PivotTable report if the report is based on an OLAP data source. The default value is True.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|vacatedStyle|string|Returns or sets the style applied to cells vacated when the PivotTable report is refreshed. The default value is a null string (no style is applied by default).|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|value|string|Returns or sets a String value that represents the name of the PivotTable report.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|version|string|Returns a value that represents the Microsoft Excel version number. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|calculatedFields|[CalculatedFieldCollection](calculatedfieldcollection.md)|Returns a CalculatedFields collection that represents all the calculated fields in the specified PivotTable report. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|dataBodyRange|[Range](range.md)|Returns a Range object that represents the range of values in a PivotTable. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|dataLabelRange|[Range](range.md)|Returns a Range object that represents the range that contains the labels for the data fields in the PivotTable report. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|pivotFields|[PivotFieldCollection](pivotfieldcollection.md)|Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of both the visible and hidden fields (a PivotFields object) in the PivotTable report. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|refreshDate|[DateTime](datetime.md)|Returns the date on which the PivotTable report was last refreshed. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[addChart(chartType: string, seriesBy: string)](#addchartcharttype-string-seriesby-string)|[Chart](chart.md)|Creates a new chart.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[addDataField(field: PivotField, caption: string, func: string)](#adddatafieldfield-pivotfield-caption-string-func-string)|[PivotField](pivotfield.md)|Adds a data field to a PivotTable report. Returns a PivotField object that represents the new data field.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[clearTable()](#cleartable)|void|The ClearTable method is used for clearing a PivotTable. Clearing PivotTables includes removing all the fields and deleting all filtering and sorting applied to the PivotTables. This method resets the PivotTable to the state it had right after it was created, before any fields were added to it.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumnField(Index: number)](#getcolumnfieldindex-number)|[PivotField](pivotfield.md)|Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently shown as column fields.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumnRange()](#getcolumnrange)|[Range](range.md)|Returns a Range object that represents the range that contains the column area in the PivotTable report.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataField(Index: number)](#getdatafieldindex-number)|[PivotField](pivotfield.md)|Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently shown as data fields.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getEntireRange()](#getentirerange)|[Range](range.md)|Returns a Range object that represents the range containing the entire PivotTable report, including page fields.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getHiddenField(index: number)](#gethiddenfieldindex-number)|[PivotField](pivotfield.md)|Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently not shown as row, column, page, or data fields.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowField(index: number)](#getrowfieldindex-number)|[PivotField](pivotfield.md)|Returns an object that represents either a single field in a PivotTable report (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently showing as row fields.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowRange()](#getrowrange)|[Range](range.md)|Returns a Range object that represents the range including the row area on the PivotTable report.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[listFormulas()](#listformulas)|void|Creates a list of calculated PivotTable items and fields on a separate worksheet.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[refreshTable()](#refreshtable)|[bool?](bool?.md)|Refreshes the PivotTable report from the source data. Returns True if it╬ô├ç├ûs successful.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[update()](#update)|void|Updates the PivotTable report.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### addChart(chartType: string, seriesBy: string)
Creates a new chart.

#### Syntax
```js
pivotTableExtObject.addChart(chartType, seriesBy);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|chartType|string|Specifies the type of chart.  Possible values are: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc.|
|seriesBy|string|Optional. Specifies the way columns or rows are used as data series on the chart.  Possible values are: Auto, Columns, Rows|

#### Returns
[Chart](chart.md)

### addDataField(field: PivotField, caption: string, func: string)
Adds a data field to a PivotTable report. Returns a PivotField object that represents the new data field.

#### Syntax
```js
pivotTableExtObject.addDataField(field, caption, func);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|field|PivotField|The unique field on the server. If the source data is Online Analytical Processing (OLAP), the unique field is a cube field. If the source data is non-OLAP (non-OLAP source data), the unique field is a PivotTable field.|
|caption|string|The label used in the PivotTable report to identify this data field.|
|func|string|The function performed in the added data field.|

#### Returns
[PivotField](pivotfield.md)

### clearTable()
The ClearTable method is used for clearing a PivotTable. Clearing PivotTables includes removing all the fields and deleting all filtering and sorting applied to the PivotTables. This method resets the PivotTable to the state it had right after it was created, before any fields were added to it.

#### Syntax
```js
pivotTableExtObject.clearTable();
```

#### Parameters
None

#### Returns
void

### getColumnField(Index: number)
Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently shown as column fields.

#### Syntax
```js
pivotTableExtObject.getColumnField(Index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|Index|number|Index of the coulmn field to be retrieved.|

#### Returns
[PivotField](pivotfield.md)

### getColumnRange()
Returns a Range object that represents the range that contains the column area in the PivotTable report.

#### Syntax
```js
pivotTableExtObject.getColumnRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getDataField(Index: number)
Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently shown as data fields.

#### Syntax
```js
pivotTableExtObject.getDataField(Index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|Index|number|Index.|

#### Returns
[PivotField](pivotfield.md)

### getEntireRange()
Returns a Range object that represents the range containing the entire PivotTable report, including page fields.

#### Syntax
```js
pivotTableExtObject.getEntireRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getHiddenField(index: number)
Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently not shown as row, column, page, or data fields.

#### Syntax
```js
pivotTableExtObject.getHiddenField(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number||

#### Returns
[PivotField](pivotfield.md)

### getRowField(index: number)
Returns an object that represents either a single field in a PivotTable report (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently showing as row fields.

#### Syntax
```js
pivotTableExtObject.getRowField(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number||

#### Returns
[PivotField](pivotfield.md)

### getRowRange()
Returns a Range object that represents the range including the row area on the PivotTable report.

#### Syntax
```js
pivotTableExtObject.getRowRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### listFormulas()
Creates a list of calculated PivotTable items and fields on a separate worksheet.

#### Syntax
```js
pivotTableExtObject.listFormulas();
```

#### Parameters
None

#### Returns
void

### refreshTable()
Refreshes the PivotTable report from the source data. Returns True if it╬ô├ç├ûs successful.

#### Syntax
```js
pivotTableExtObject.refreshTable();
```

#### Parameters
None

#### Returns
[bool?](bool?.md)

### update()
Updates the PivotTable report.

#### Syntax
```js
pivotTableExtObject.update();
```

#### Parameters
None

#### Returns
void
