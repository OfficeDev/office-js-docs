# NamedItemCollection Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

A collection of all the nameditem objects that are part of the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[NamedItem[]](nameditem.md)|A collection of namedItem objects. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|Gets a nameditem object using its name|1.1|
|[getItemOrNull(name: string)](#getitemornullname-string)|[NamedItem](nameditem.md)|Gets a nameditem object using its name. If the nameditem object does not exist, the returned object's isNull property will be true.|1.3|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### getItem(name: string)
Gets a nameditem object using its name

#### Syntax
```js
namedItemCollectionObject.getItem(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|name|string|nameditem name.|

#### Returns
[NamedItem](nameditem.md)

### getItemOrNull(name: string)
Gets a nameditem object using its name. If the nameditem object does not exist, the returned object's isNull property will be true.

#### Syntax
```js
namedItemCollectionObject.getItemOrNull(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|name|string|nameditem name.|

#### Returns
[NamedItem](nameditem.md)

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
