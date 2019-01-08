# PivotLayout Object (JavaScript API for Excel)

Represents the visual layout of the PivotTable.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|autoFormat|bool|True if formatting will be automatically formatted when it╬ô├ç├ûs refreshed or when fields are moved|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|enableFieldList|bool|True if the field list should be shown or hidden from the UI.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|preserveFormatting|bool|True if formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|showColumnGrandTotals|bool|True if the PivotTable report shows grand totals for columns.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|showRowGrandTotals|bool|True if the PivotTable report shows grand totals for rows.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|subtotalLocation|string|This property indicates the SubtotalLocationType of all fields on the PivotTable. If fields have different states, this will be null. Possible values are: AtTop, AtBottom.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|layoutType|[PivotLayoutType](pivotlayouttype.md)|This property indicates the PivotLayoutType of all fields on the PivotTable. If fields have different states, this will be null.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCell(dataHierarchy: DataPivotHierarchy or string, rowItems: ()[], columnItems: ()[])](#getcelldatahierarchy-datapivothierarchy-or-string-rowitems--columnitems-)|[Range](range.md)|Gets the cell in the PivotTable's data body that contains the value for the intersection of the specified dataHierarchy, rowItems, and columnItems.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumnLabelRange()](#getcolumnlabelrange)|[Range](range.md)|Returns the range where the PivotTable's column labels reside.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Returns the range where the PivotTable's data values reside.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataHierarchy(cell: Range or string)](#getdatahierarchycell-range-or-string)|[DataPivotHierarchy](datapivothierarchy.md)|Gets the DataHierarchy that is used to calculate the value in a specified range within the PivotTable.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getFilterAxisRange()](#getfilteraxisrange)|[Range](range.md)|Returns the range of the PivotTable's filter area.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getPivotItems(axis: PivotAxis, cell: Range or string)](#getpivotitemsaxis-pivotaxis-cell-range-or-string)|[PivotItem[]](pivotitem[].md)|Gets the PivotItems from an axis that make up the value in a specified range within the PivotTable.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|Returns the range the PivotTable exists on, excluding the filter area.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowLabelRange()](#getrowlabelrange)|[Range](range.md)|Returns the range where the PivotTable's row labels reside.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[setAutosortOnCell(cell: Range or string, sortby: SortBy)](#setautosortoncellcell-range-or-string-sortby-sortby)|void|Sets an autosort using the specified cell to automatically select all criteria and context for the sort.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCell(dataHierarchy: DataPivotHierarchy or string, rowItems: ()[], columnItems: ()[])
Gets the cell in the PivotTable's data body that contains the value for the intersection of the specified dataHierarchy, rowItems, and columnItems.

#### Syntax
```js
pivotLayoutObject.getCell(dataHierarchy, rowItems, columnItems);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|dataHierarchy|DataPivotHierarchy or string|The dataHierarchy that provides the data item to find.|
|rowItems|()[]|The PivotItems from the row axis that make up the value to find.|
|columnItems|()[]|The PivotItems from the column axis that make up the value to find.|

#### Returns
[Range](range.md)

### getColumnLabelRange()
Returns the range where the PivotTable's column labels reside.

#### Syntax
```js
pivotLayoutObject.getColumnLabelRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getDataBodyRange()
Returns the range where the PivotTable's data values reside.

#### Syntax
```js
pivotLayoutObject.getDataBodyRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getDataHierarchy(cell: Range or string)
Gets the DataHierarchy that is used to calculate the value in a specified range within the PivotTable.

#### Syntax
```js
pivotLayoutObject.getDataHierarchy(cell);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|cell|Range or string|A single cell within the PivotTable data body to get the data hieararchy for.|

#### Returns
[DataPivotHierarchy](datapivothierarchy.md)

### getFilterAxisRange()
Returns the range of the PivotTable's filter area.

#### Syntax
```js
pivotLayoutObject.getFilterAxisRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getPivotItems(axis: PivotAxis, cell: Range or string)
Gets the PivotItems from an axis that make up the value in a specified range within the PivotTable.

#### Syntax
```js
pivotLayoutObject.getPivotItems(axis, cell);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|axis|PivotAxis|The axis to get the PivotItems from. Must be either "row" or "column."|
|cell|Range or string|A single cell within the PivotTable's data body to get the PivotItems for.|

#### Returns
[PivotItem[]](pivotitem[].md)

### getRange()
Returns the range the PivotTable exists on, excluding the filter area.

#### Syntax
```js
pivotLayoutObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getRowLabelRange()
Returns the range where the PivotTable's row labels reside.

#### Syntax
```js
pivotLayoutObject.getRowLabelRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### setAutosortOnCell(cell: Range or string, sortby: SortBy)
Sets an autosort using the specified cell to automatically select all criteria and context for the sort.

#### Syntax
```js
pivotLayoutObject.setAutosortOnCell(cell, sortby);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|cell|Range or string|A single cell to use get the criteria from for applying the autosort.|
|sortby|SortBy|The direction of the sort.|

#### Returns
void
