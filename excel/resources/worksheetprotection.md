# WorksheetProtection Object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

Represents the protection of a sheet object.

_**Note**: This is a proposed feature that is still under design phase and hence not yet available as part of the product. The specification is being made available for community review and feedback. The final design may change. Help us make this feature better by providing your feedback [here](https://github.com/OfficeDev/office-js-docs/issues/new?title=ExcelJs-1.2-OpenSpec-worksheetprotection)._

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|protected|bool|Indicates if the worksheet is protected. Read-Only. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|options|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|Sheet protection options. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[protect(options: WorksheetProtectionOptions, password: string)](#protectoptions-worksheetprotectionoptions-password-string)|void|Protect a worksheet. It throws if the worksheet has been protected.|
|[unprotect(password: string)](#unprotectpassword-string)|void|Unprotect a worksheet|

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

### protect(options: WorksheetProtectionOptions, password: string)
Protect a worksheet. It throws if the worksheet has been protected.

#### Syntax
```js
worksheetProtectionObject.protect(options, password);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|options|WorksheetProtectionOptions|Optional. sheet protection options.|
|password|string|Optional. sheet protection password.|

#### Returns
void

### unprotect(password: string)
Unprotect a worksheet

#### Syntax
```js
worksheetProtectionObject.unprotect(password);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|password|string|Optional. sheet protection password.|

#### Returns
void
