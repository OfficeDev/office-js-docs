# ConditionalFormatColorScale Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents the conditional format for a Color Scale.

## Properties

None

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|maximum|[ConditionalFormatColorScaleRule](conditionalformatcolorscalerule.md)|Represents the maximum rule and color for the Color Scale.|1.3||
|midpoint|[ConditionalFormatColorScaleRule](conditionalformatcolorscalerule.md)|Represents the mid-point rule and color for the Color Scale.|1.3||
|minimum|[ConditionalFormatColorScaleRule](conditionalformatcolorscalerule.md)|Represents the minimum rule and color for the Color Scale.|1.3||

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
