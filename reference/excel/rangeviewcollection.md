# RangeViewCollection Object (JavaScript API for Excel)

Represents a collection of RangeView objects.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[RangeView[]](rangeview.md)|A collection of rangeView objects. Read-only.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Gets the number of RangeView objects in the collection.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getFirst()](#getfirst)|[RangeView](rangeview.md)|Gets the first RangeView object in the collection.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[RangeView](rangeview.md)|Gets a RangeView Row via it's index. Zero-Indexed.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getLast()](#getlast)|[RangeView](rangeview.md)|Gets the last RangeView object in the collection.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCount()
Gets the number of RangeView objects in the collection.

#### Syntax
```js
rangeViewCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getFirst()
Gets the first RangeView object in the collection.

#### Syntax
```js
rangeViewCollectionObject.getFirst();
```

#### Parameters
None

#### Returns
[RangeView](rangeview.md)

### getItemAt(index: number)
Gets a RangeView Row via it's index. Zero-Indexed.

#### Syntax
```js
rangeViewCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|Index of the visible row.|

#### Returns
[RangeView](rangeview.md)

### getLast()
Gets the last RangeView object in the collection.

#### Syntax
```js
rangeViewCollectionObject.getLast();
```

#### Parameters
None

#### Returns
[RangeView](rangeview.md)

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
