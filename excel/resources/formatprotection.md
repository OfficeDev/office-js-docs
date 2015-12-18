# FormatProtection Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

Represents the format protection of a range object.

_**Note**: This is a proposed feature that is still under design phase and hence not yet available as part of the product. The specification is being made available for community review and feedback. The final design may change. Help us make this feature better by providing your feedback [here](https://github.com/OfficeDev/office-js-docs/issues/new?title=ExcelJs-1.2-OpenSpec-worksheetprotection)._

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|formulaHidden|bool|Indicates if Excel hides the formula for the cells in the range. A null value indicates that the entire range doesn't have uniform formula hidden setting.|
|locked|bool|Indicates if Excel locks the cells in the object. A null value indicates that the entire range doesn't have uniform lock setting.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
