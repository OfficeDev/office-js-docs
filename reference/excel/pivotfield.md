# PivotField Object (JavaScript API for Excel)

Represents a field in a PivotTable report.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|aggregationFunction|string|Returns or sets the function used to summarize the PivotTable field (data fields only).|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|allItemsVisible|bool|Used to retrieve a Boolean value that indicates whether or not any manual filtering is applied to the PivotField. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|autoSortField|string|Returns the name of the data field used to sort the specified PivotTable field automatically. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|autoSortOrder|string|Returns the order used to sort the specified PivotTable field automatically. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|calculated|bool|True if the PivotTable field is a calculated field or item. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|calculation|string|Returns or sets a PivotFieldCalculation value that represents the type of calculation performed by the specified field. This property is valid only for data fields.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|caption|string|Returns a String value that represents the label text for the pivot field.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|dataType|string|Returns a PivotFieldDataType value that represents the type of data in the PivotTable field. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|drilledDown|bool|True if the flag for the specified PivotTable field or PivotTable item is set to drilled (expanded, or visible).|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|enableMultiplePageItems|bool|Used for specifying whether or not check boxes are present in the filter drop-down list for fields in the page area.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formula|string|Returns or sets a string value that represents the object's formula in A1-style notation.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Returns or sets a String value representing the name of the object.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|string|Returns or sets a String value that represents the format code for the object.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|orientation|string|Returns or sets a value that represents the location of the field in the specified PivotTable report.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|position|int|Returns or sets a value that represents the position of the field (first, second, third, and so on) among all the fields in its orientation (Rows, Columns, Pages, Data).|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showDetail|bool|Gets or sets whether the specified PivotField is showing detail.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|sourceName|string|Returns a string value that represents the specified object╬ô├ç├ûs name as it appears in the original source data for the specified PivotTable report. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|currentPage|[PivotItem](pivotitem.md)|Returns or sets the current page showing for the page field (valid only for page fields).|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|hiddenItems|[PivotItemCollection](pivotitemcollection.md)|Returns an object that represents either a single hidden PivotTable item (a PivotItemobject) or a collection of all the hidden items (a PivotItems object) in the specified field. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|pivotItems|[PivotItemCollection](pivotitemcollection.md)|Returns an object that represents either a single PivotTable item (a PivotItem object) or a collection of all the visible and hidden items (a PivotItems object) in the specified field. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|subtotals|[Subtotals](subtotals.md)|Returns or sets subtotals displayed with the specified field. Valid only for nondata fields.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visiblePivotItems|[PivotItemCollection](pivotitemcollection.md)|Returns an object that represents either a single visible PivotTable item (a PivotItemobject) or a collection of all the visible items (a PivotItems object) in the specified field. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[autoGroup()](#autogroup)|void|Automatically groups the pivot fields in a pivot table.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[autoSort(sortOrder: string, Field: string)](#autosortsortorder-string-field-string)|void|Establishes automatic field-sorting rules for PivotTable reports.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[clearAllFilters()](#clearallfilters)|void|Calling this method deletes all filters currently applied to the PivotField. This includes deleting all filters from the PivotFilters collection of the PivotField and removing any manual filtering applied to the PivotField as well. If the PivotField is in the Report Filter area, the item selected will be set to the default item.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataRange()](#getdatarange)|[Range](range.md)|Returns a Range object that represents the range of the current PivotField.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### autoGroup()
Automatically groups the pivot fields in a pivot table.

#### Syntax
```js
pivotFieldObject.autoGroup();
```

#### Parameters
None

#### Returns
void

### autoSort(sortOrder: string, Field: string)
Establishes automatic field-sorting rules for PivotTable reports.

#### Syntax
```js
pivotFieldObject.autoSort(sortOrder, Field);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|sortOrder|string|Specifies the sort order.|
|Field|string|The name of the sort key field. You must specify the unique name, and not the displayed name.|

#### Returns
void

### clearAllFilters()
Calling this method deletes all filters currently applied to the PivotField. This includes deleting all filters from the PivotFilters collection of the PivotField and removing any manual filtering applied to the PivotField as well. If the PivotField is in the Report Filter area, the item selected will be set to the default item.

#### Syntax
```js
pivotFieldObject.clearAllFilters();
```

#### Parameters
None

#### Returns
void

### getDataRange()
Returns a Range object that represents the range of the current PivotField.

#### Syntax
```js
pivotFieldObject.getDataRange();
```

#### Parameters
None

#### Returns
[Range](range.md)
