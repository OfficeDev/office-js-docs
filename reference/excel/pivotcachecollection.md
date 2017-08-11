# PivotCacheCollection Object (JavaScript API for Excel)

Represents the collection of memory caches from the PivotTable reports in a workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[PivotCache[]](pivotcache.md)|A collection of pivotCache objects. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(sourceType: string, address: Range)](#addsourcetype-string-address-range)|[PivotCache](pivotcache.md)|Creates a new PivotCache.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Returns a value that represents the number of objects in the collection.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(index: number)](#getitemindex-number)|[PivotCache](pivotcache.md)|Returns a single object from a collection based on the index.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[PivotCache](pivotcache.md)|Returns a single object from a collection based on the index.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(sourceType: string, address: Range)
Creates a new PivotCache.

#### Syntax
```js
pivotCacheCollectionObject.add(sourceType, address);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|sourceType|string|Determines the type of PivotCache.|
|address|Range|The data for the new PivotTable cache.|

#### Returns
[PivotCache](pivotcache.md)

### getCount()
Returns a value that represents the number of objects in the collection.

#### Syntax
```js
pivotCacheCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(index: number)
Returns a single object from a collection based on the index.

#### Syntax
```js
pivotCacheCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index of the item to be retrieved.|

#### Returns
[PivotCache](pivotcache.md)

### getItemAt(index: number)
Returns a single object from a collection based on the index.

#### Syntax
```js
pivotCacheCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index of the item. Zero-indexed.|

#### Returns
[PivotCache](pivotcache.md)
