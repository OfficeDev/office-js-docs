# ConditionalIconCriterion Object (JavaScript API for Excel)

Represents an Icon Criterion which contains a type, value, an Operator, and an optional custom icon, if not using an iconset.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|formula|object|A number or a formula depending on the type.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|customIcon|[Icon](icon.md)|The custom icon for the current criterion if different from the default IconSet, else null will be returned.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|operator|[ConditionalIconCriterionOperator](conditionaliconcriterionoperator.md)|GreaterThan or GreaterThanOrEqual for each of the rule type for the Icon conditional format.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|type|[ConditionalFormatIconRuleType](conditionalformaticonruletype.md)|What the icon conditional formula should be based on.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

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
