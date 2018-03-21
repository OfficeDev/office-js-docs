# AreaCollection Object (JavaScript API for Excel)

Represents a collection of all the Data Connections that are part of the workbook or worksheet.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[Area[]](area.md)|A collection of area objects. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Gets the number of contiguous areas in a range.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Range](range.md)|Gets a contiguous area object based on its position in the collection.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCount()
Gets the number of contiguous areas in a range.

#### Syntax
```js
areaCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItemAt(index: number)
Gets a contiguous area object based on its position in the collection.

#### Syntax
```js
areaCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[Range](range.md)
