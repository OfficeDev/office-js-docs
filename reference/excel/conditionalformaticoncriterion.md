# ConditionalFormatIconCriterion Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents an Icon Criterion which contains a type, value, an Operator, and an optional custom icon, if not using an iconset.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|formula|object|A number, a formula, or null (if Type is LowestValue, HighestValue, or Automatic).|1.3||
|type|string|What the icon conditional formula should be based on. Possible values are: LowestValue, HighestValue, Number, Percent, Formula, Percentile, Automatic.|1.3||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|customIcon|[Icon](icon.md)|An icon type, where one can select a specific icon for this criterion.|1.3||
|operator|[IconCriterionOperator](iconcriterionoperator.md)|GreaterThan or GreaterThanOrEqual for each of the rule type for the Icon conditional format.|1.3||

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
