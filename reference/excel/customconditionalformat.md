# CustomConditionalFormat Object (JavaScript API for Excel)

Represents a custom conditional format type.

## Properties

None

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|format|[ConditionalRangeFormat](conditionalrangeformat.md)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|rule|[ConditionalFormatRule](conditionalformatrule.md)|Represents the Rule object on this conditional format. Read-only.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

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
