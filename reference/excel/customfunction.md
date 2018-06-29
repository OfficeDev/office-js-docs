# CustomFunction Object (JavaScript API for Excel)

Custom function declaration.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|batching|bool|Represents whether the function expects multiple argument sets at a time. Read-only.|[ApiSet.CustomBase.CustomFunctions](../requirement-sets/excel-api-requirement-sets.md)|
|cancelable|bool|Represents whether the function supports cancellation. Read-only.|[ApiSet.CustomBase.CustomFunctions](../requirement-sets/excel-api-requirement-sets.md)|
|description|string|A description of the custom function. Read-only.|[ApiSet.CustomBase.CustomFunctions](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|The ID of the function. Read-only.|[ApiSet.CustomBase.CustomFunctions](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|The name of the function. Read-only.|[ApiSet.CustomBase.CustomFunctions](../requirement-sets/excel-api-requirement-sets.md)|
|streaming|bool|Represents whether the function supports returning results multiple times. Read-only.|[ApiSet.CustomBase.CustomFunctions](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|parameters|[CustomFunctionParameter](customfunctionparameter.md)|Parameters of the function. Read-only.|[ApiSet.CustomBase.CustomFunctions](../requirement-sets/excel-api-requirement-sets.md)|
|resultDimensionality|[CustomFunctionDimensionality](customfunctiondimensionality.md)|The dimensionality of result values. Read-only.|[ApiSet.CustomBase.CustomFunctions](../requirement-sets/excel-api-requirement-sets.md)|
|resultType|[CustomFunctionValueType](customfunctionvaluetype.md)|Result type returned by the function. Read-only.|[ApiSet.CustomBase.CustomFunctions](../requirement-sets/excel-api-requirement-sets.md)|
|type|[CustomFunctionType](customfunctiontype.md)|The type of the function. Read-only.|[ApiSet.CustomBase.CustomFunctions](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes this function from Excel.|[ApiSet.CustomBase.CustomFunctions](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### delete()
Deletes this function from Excel.

#### Syntax
```js
customFunctionObject.delete();
```

#### Parameters
None

#### Returns
void
