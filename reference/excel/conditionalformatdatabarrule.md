# ConditionalFormatDataBarRule Object (JavaScript API for Excel)

Represents a rule-type for a Data Bar.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|formula|object|The formula, if required, to evaluate the databar rule on.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|formulaLocal|object|The formula, if required, to evaluate the databar rule on in the user's language.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|formulaR1C1|object|The formula, if required, to evaluate the databar rule on in R1C1-style notation.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|What the databar should be based on. Possible values are: LowestValue, HighestValue, Number, Percent, Formula, Percentile, Automatic.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

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
