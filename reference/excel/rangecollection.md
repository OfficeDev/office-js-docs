# RangeCollection Object (JavaScript API for Excel)

Represents a collection of all the Data Connections that are part of the workbook or worksheet.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[Range[]](range.md)|A collection of range objects. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Returns the number of ranges in the RangeCollection.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Range](range.md)|Returns the range object based on its position in the RangeCollection.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCount()
Returns the number of ranges in the RangeCollection.

#### Syntax
```js
rangeCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItemAt(index: number)
Returns the range object based on its position in the RangeCollection.

#### Syntax
```js
rangeCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the range object to be retrieved. Zero-indexed.|

#### Returns
[Range](range.md)
