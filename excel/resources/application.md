# Application Object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

Represents the Excel application that manages the workbook.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|calculationMode|string|Returns the calculation mode used in the workbook. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|Recalculate all currently opened workbooks in Excel.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### calculate(calculationType: string)
Recalculate all currently opened workbooks in Excel.

#### Syntax
```js
applicationObject.calculate(calculationType);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|calculationType|string|Specifies the calculation type to use.|

#### Returns
void

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
