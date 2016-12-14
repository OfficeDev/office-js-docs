# ConditionalDataBarNegativeFormat Object (JavaScript API for Excel)

Represents a conditional format DataBar Format for the negative side of the data bar.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|borderColor|string|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|fillColor|string|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|matchPositiveBorderColor|bool|Boolean representation of whether or not the negative DataBar has the same border color as the positive DataBar.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|matchPositiveFillColor|bool|Boolean representation of whether or not the negative DataBar has the same fill color as the positive DataBar.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


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
