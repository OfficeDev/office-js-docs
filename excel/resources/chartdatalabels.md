# ChartDataLabels Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents a collection of all the data labels on a chart point.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|separator|string|String representing the separator used for the data labels on a chart.|1.1||
|showBubbleSize|bool|Boolean value representing if the data label bubble size is visible or not.|1.1||
|showCategoryName|bool|Boolean value representing if the data label category name is visible or not.|1.1||
|showLegendKey|bool|Boolean value representing if the data label legend key is visible or not.|1.1||
|showPercentage|bool|Boolean value representing if the data label percentage is visible or not.|1.1||
|showSeriesName|bool|Boolean value representing if the data label series name is visible or not.|1.1||
|showValue|bool|Boolean value representing if the data label value is visible or not.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartDataLabelFormat](chartdatalabelformat.md)|Represents the format of chart data labels, which includes fill and font formatting. Read-only.|1.1||
|position|string|DataLabelPosition value that represents the position of the data label.|1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


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
