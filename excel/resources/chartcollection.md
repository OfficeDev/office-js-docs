# ChartCollection Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

A collection of all the chart objects on a worksheet.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|count|int|Returns the number of charts in the worksheet. Read-only.|1.1||
|items|[Chart[]](chart.md)|A collection of chart objects. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(type: string, sourceData: Range, seriesBy: string)](#addtype-string-sourcedata-range-seriesby-string)|[Chart](chart.md)|Creates a new chart.|1.1|
|[getItem(name: string)](#getitemname-string)|[Chart](chart.md)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|1.1|
|[getItemAt(index: number)](#getitematindex-number)|[Chart](chart.md)|Gets a chart based on its position in the collection.|1.1|
|[getItemOrNull(name: string)](#getitemornullname-string)|[Chart](chart.md)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|1.3|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### add(type: string, sourceData: Range, seriesBy: string)
Creates a new chart.

#### Syntax
```js
chartCollectionObject.add(type, sourceData, seriesBy);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|type|string|Represents the type of a chart.|
|sourceData|Range|The Range object corresponding to the source data.|
|seriesBy|string|Optional. Specifies the way columns or rows are used as data series on the chart.|

#### Returns
[Chart](chart.md)

### getItem(name: string)
Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.

#### Syntax
```js
chartCollectionObject.getItem(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|name|string|Name of the chart to be retrieved.|

#### Returns
[Chart](chart.md)

### getItemAt(index: number)
Gets a chart based on its position in the collection.

#### Syntax
```js
chartCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[Chart](chart.md)

### getItemOrNull(name: string)
Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.

#### Syntax
```js
chartCollectionObject.getItemOrNull(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|name|string|Name of the chart to be retrieved.|

#### Returns
[Chart](chart.md)

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
