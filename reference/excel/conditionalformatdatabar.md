# ConditionalFormatDataBar Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents an Excel Conditional Data Bar Type.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|axisFormat|string|Representation of how the axis is determined for an Excel data bar. Possible values are: Automatic, None, CellMidPoint.|1.3||
|direction|string|Represents the direction that the data bar graphic should be based on. Possible values are: LeftToRight, Context, RightToLeft.|1.3||
|showDataBarOnly|bool|If true, hides the values from the cells where the data bar is applied.|1.3||

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|lowerBoundRule|[ConditionalFormatDataBarRule](conditionalformatdatabarrule.md)|The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.|1.3||
|negativeFormat|[ConditionalFormatDataBarFormat](conditionalformatdatabarformat.md)|Representation of all values to the left of the axis in an Excel data bar.|1.3||
|positiveFormat|[ConditionalFormatDataBarFormat](conditionalformatdatabarformat.md)|Representation of all values to the right of the axis in an Excel data bar.|1.3||
|upperBoundRule|[ConditionalFormatDataBarRule](conditionalformatdatabarrule.md)|The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.|1.3||

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
