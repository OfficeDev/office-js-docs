# SlicerCollection Object (JavaScript API for Excel)

Represents a collection of all the slicer objects on the workbook or a worksheet.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[Slicer[]](slicer.md)|A collection of slicer objects. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(slicerSource: object, sourceField: object, slicerDestination: object)](#addslicersource-object-sourcefield-object-slicerdestination-object)|[Slicer](slicer.md)|Adds a new slicer to the workbook.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Returns the number of slicers in the collection.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[Slicer](slicer.md)|Gets a slicer object using its name or id.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Slicer](slicer.md)|Gets a slicer based on its position in the collection.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: string)](#getitemornullobjectkey-string)|[Slicer](slicer.md)|Gets a slicer using its name or id. If the slicer does not exist, will return a null object.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(slicerSource: object, sourceField: object, slicerDestination: object)
Adds a new slicer to the workbook.

#### Syntax
```js
slicerCollectionObject.add(slicerSource, sourceField, slicerDestination);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|slicerSource|object|The data source that the new slicer will be based on. It can be a PivotTable object, a Table object or a string. When a PivotTable object is passed, the data source is the source of the PivotTable object. When a Table object is passed, the data source is the Table object. When a string is passed, it is interpreted as the name/id of a PivotTable/Table.|
|sourceField|object|The field in the data source to filter by. It can be a PivotField object, a TableColumn object, the id of a PivotField or the id/name of TableColumn.|
|slicerDestination|object|Optional. The worksheet where the new slicer will be created in. It can be a Worksheet object or the name/id of a worksheet. This parameter can be omitted if the slicer collection is retrieved from worksheet.|

#### Returns
[Slicer](slicer.md)

### getCount()
Returns the number of slicers in the collection.

#### Syntax
```js
slicerCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(key: string)
Gets a slicer object using its name or id.

#### Syntax
```js
slicerCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|string|The name or id of the slicer.|

#### Returns
[Slicer](slicer.md)

### getItemAt(index: number)
Gets a slicer based on its position in the collection.

#### Syntax
```js
slicerCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|index|

#### Returns
[Slicer](slicer.md)

### getItemOrNullObject(key: string)
Gets a slicer using its name or id. If the slicer does not exist, will return a null object.

#### Syntax
```js
slicerCollectionObject.getItemOrNullObject(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|string|Name or Id of the slicer to be retrieved.|

#### Returns
[Slicer](slicer.md)
