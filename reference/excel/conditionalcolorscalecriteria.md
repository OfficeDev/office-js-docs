# ConditionalColorScaleCriteria Object (JavaScript API for Excel)

Represents the criteria of the color scale.

## Properties

None

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|maximum|[ConditionalColorScaleCriterion](conditionalcolorscalecriterion.md)|The maximum point Color Scale Criterion.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|midpoint|[ConditionalColorScaleCriterion](conditionalcolorscalecriterion.md)|The midpoint Color Scale Criterion if the color scale is a 3-color scale.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|minimum|[ConditionalColorScaleCriterion](conditionalcolorscalecriterion.md)|The minimum point Color Scale Criterion.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

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
