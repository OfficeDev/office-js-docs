# RangeBorderCollection Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents the border objects that make up range border.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|count|int|Number of border objects in the collection. Read-only.|1.1||
|items|[RangeBorder[]](rangeborder.md)|A collection of rangeBorder objects. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getItem(index: string)](#getitemindex-string)|[RangeBorder](rangeborder.md)|Gets a border object using its name|1.1|
|[getItemAt(index: number)](#getitematindex-number)|[RangeBorder](rangeborder.md)|Gets a border object using its index|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### getItem(index: string)
Gets a border object using its name

#### Syntax
```js
rangeBorderCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|string|Index value of the border object to be retrieved.|

#### Returns
[RangeBorder](rangeborder.md)

### getItemAt(index: number)
Gets a border object using its index

#### Syntax
```js
rangeBorderCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[RangeBorder](rangeborder.md)

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
