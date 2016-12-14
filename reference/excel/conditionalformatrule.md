# ConditionalFormatRule Object (JavaScript API for Excel)

Represents a rule, for all traditional ruleformat pairings.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|formula1|object|The formula, if required, to evaluate the conditional format rule on.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|formula1Local|object|The formula, if required, to evaluate the conditional format rule on in the user's language.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|formula1R1C1|object|The formula, if required, to evaluate the conditional format rule on in R1C1-style notation.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|formula2|object|The formula, if required, to evaluate the conditional format rule on.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|formula2Local|object|The formula, if required, to evaluate the conditional format rule on in the user's language.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|formula2R1C1|object|The formula, if required, to evaluate the conditional format rule on in R1C1-style notation.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

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
