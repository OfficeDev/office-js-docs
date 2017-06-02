# ChartTrendlineCollection Object (JavaScript API for Excel)

Represents a collection of Chart Trendlines.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[ChartTrendline[]](charttrendline.md)|A collection of chartTrendline objects. Read-only.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(type: string)](#addtype-string)|[ChartTrendline](charttrendline.md)|Adds a new trendline to trendline collection.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Returns the number of trendlines in the collection.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(index: int)](#getitemindex-int)|[ChartTrendline](charttrendline.md)|Get trendline object by index, which is the insertion order in items array.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|

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
|type|string|Optional. Specifies the trendline type.  Possible values are: Linear, Expontential, Logarithmic, MovingAvg, Polynomial, Power|

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

### getItem(index: int)
Get trendline object by index, which is the insertion order in items array.

#### Syntax
```js
chartTrendlineCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|int|Represents the insertion order in items array.|

#### Returns
[ChartTrendline](charttrendline.md)

