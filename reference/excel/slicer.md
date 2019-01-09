# Slicer Object (JavaScript API for Excel)

Represents a slicer object in the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|caption|string|Represents the caption of slicer.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|height|double|Represents the height, in points, of the slicer.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|Represents the unique id of slicer. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|isFilterCleared|bool|True if all filters currently applied on the slicer is cleared. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|left|double|Represents the distance, in points, from the left side of the slicer to the left of the worksheet.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Represents the name of slicer.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|nameInFormula|string|Represents the name used in the formula.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|style|string|Constant value that represents the Slicer style. Possible values are: SlicerStyleLight1 thru SlicerStyleLight6, TableStyleOther1 thru TableStyleOther2, SlicerStyleDark1 thru SlicerStyleDark6. A custom user-defined style present in the workbook can also be specified.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|top|double|Represents the distance, in points, from the top edge of the slicer to the right of the worksheet.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|width|double|Represents the width, in points, of the slicer.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|slicerItems|[SlicerItemCollection](sliceritemcollection.md)|Represents the collection of SlicerItems that are part of the slicer. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|sortBy|[SlicerSortType](slicersorttype.md)|Represents the sort order of the items in the slicer.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|worksheet|[Worksheet](worksheet.md)|Represents the worksheet containing the slicer. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clearFilters()](#clearfilters)|void|Clears all the filters currently applied on the slicer.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[delete()](#delete)|void|Deletes the slicer.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getSelectedItems()](#getselecteditems)|[string[]](string[].md)|Returns an array of selected items' names. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[selectItems(items: string[])](#selectitemsitems-string)|void|Select slicer items based on their names. Previous selection will be cleared.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### clearFilters()
Clears all the filters currently applied on the slicer.

#### Syntax
```js
slicerObject.clearFilters();
```

#### Parameters
None

#### Returns
void

### delete()
Deletes the slicer.

#### Syntax
```js
slicerObject.delete();
```

#### Parameters
None

#### Returns
void

### getSelectedItems()
Returns an array of selected items' names. Read-only.

#### Syntax
```js
slicerObject.getSelectedItems();
```

#### Parameters
None

#### Returns
[string[]](string[].md)

### selectItems(items: string[])
Select slicer items based on their names. Previous selection will be cleared.

#### Syntax
```js
slicerObject.selectItems(items);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|items|string[]|Optional. Optional. The specified slicer item names to be selected.|

#### Returns
void
