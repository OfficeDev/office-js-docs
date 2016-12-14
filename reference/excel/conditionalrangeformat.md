# ConditionalRangeFormat Object (JavaScript API for Excel)

A format object encapsulating the conditional formats range's font, fill, borders, and other properties.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|numberFormat|object|Represents Excel's number format code for the given range. Cleared if null is passed in.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|borders|[ConditionalRangeBorderCollection](conditionalrangebordercollection.md)|Collection of border objects that apply to the overall conditional format range. Read-only.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|fill|[ConditionalRangeFill](conditionalrangefill.md)|Returns the fill object defined on the overall conditional format range. Read-only.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|font|[ConditionalRangeFont](conditionalrangefont.md)|Returns the font object defined on the overall conditional format range. Read-only.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

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
