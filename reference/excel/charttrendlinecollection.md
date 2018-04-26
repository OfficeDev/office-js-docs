# ChartTrendlineCollection Object (JavaScript API for Excel)

Represents a collection of Chart Trendlines.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[ChartTrendline[]](charttrendline.md)|A collection of chartTrendline objects. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(type: string)](#addtype-string)|[ChartTrendline](charttrendline.md)|Adds a new trendline to trendline collection.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Returns the number of trendlines in the collection.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(index: number)](#getitemindex-number)|[ChartTrendline](charttrendline.md)|Get trendline object by index, which is the insertion order in items array.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(type: string)
Adds a new trendline to trendline collection.

#### Syntax
```js
chartTrendlineCollectionObject.add(type);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|type|string|Optional. Specifies the trendline type. The default value is "Linear".|

#### Returns
[ChartTrendline](charttrendline.md)

### getCount()
Returns the number of trendlines in the collection.

#### Syntax
```js
chartTrendlineCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(index: number)
Get trendline object by index, which is the insertion order in items array.

#### Syntax
```js
chartTrendlineCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Represents the insertion order in items array.|

#### Returns
[ChartTrendline](charttrendline.md)
