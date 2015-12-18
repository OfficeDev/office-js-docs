# ChartLegend Object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

Represents the legend in a chart.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|overlay|bool|Boolean value for whether the chart legend should overlap with the main body of the chart.|
|position|string|Represents the position of the legend on the chart.|
|visible|bool|A boolean value the represents the visibility of a ChartLegend object.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|format|[ChartLegendFormat](chartlegendformat.md)|Represents the formatting of a chart legend, which includes fill and font formatting. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
