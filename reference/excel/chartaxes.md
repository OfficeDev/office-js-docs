# ChartAxes Object (JavaScript API for Excel)

Represents the chart axes.

## Properties

None

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|categoryAxis|[ChartAxis](chartaxis.md)|Represents the category axis in a chart. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|seriesAxis|[ChartAxis](chartaxis.md)|Represents the series axis of a 3-dimensional chart. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|valueAxis|[ChartAxis](chartaxis.md)|Represents the value axis in an axis. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getItem(type: string, group: string)](#getitemtype-string-group-string)|[ChartAxis](chartaxis.md)|Returns the specific axis identified by type and group.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getItem(type: string, group: string)
Returns the specific axis identified by type and group.

#### Syntax
```js
chartAxesObject.getItem(type, group);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|type|string|Specifies the axis type.|
|group|string|Optional. Specifies the axis group.|

#### Returns
[ChartAxis](chartaxis.md)
