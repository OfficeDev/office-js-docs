# DataBarConditionalFormat Object (JavaScript API for Excel)

Represents an Excel Conditional Data Bar Type.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|axisColor|string|HTML color code representing the color of the Axis line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|showDataBarOnly|bool|If true, hides the values from the cells where the data bar is applied.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|axisFormat|[ConditionalDataBarAxisFormat](conditionaldatabaraxisformat.md)|Representation of how the axis is determined for an Excel data bar.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|barDirection|[ConditionalDataBarDirection](conditionaldatabardirection.md)|Represents the direction that the data bar graphic should be based on.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|lowerBoundRule|[ConditionalDataBarRule](conditionaldatabarrule.md)|The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|negativeFormat|[ConditionalDataBarNegativeFormat](conditionaldatabarnegativeformat.md)|Representation of all values to the left of the axis in an Excel data bar. Read-only.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|positiveFormat|[ConditionalDataBarPositiveFormat](conditionaldatabarpositiveformat.md)|Representation of all values to the right of the axis in an Excel data bar. Read-only.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|upperBoundRule|[ConditionalDataBarRule](conditionaldatabarrule.md)|The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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
