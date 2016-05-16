# WorksheetProtection Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents the protection of a sheet object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|protected|bool|Indicates if the worksheet is protected. Read-Only. Read-only.|1.2||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|options|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|Sheet protection options. Read-only.|1.2||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|
|[protect(options: WorksheetProtectionOptions)](#protectoptions-worksheetprotectionoptions)|void|Protects a worksheet. Fails if the worksheet has been protected.|1.2|
|[unprotect()](#unprotect)|void|Unprotects a worksheet.|1.2|

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

### protect(options: WorksheetProtectionOptions)
Protects a worksheet. Fails if the worksheet has been protected.

#### Syntax
```js
worksheetProtectionObject.protect(options);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|options|WorksheetProtectionOptions|Optional. sheet protection options.|

#### Returns
void

### unprotect()
Unprotects a worksheet.

#### Syntax
```js
worksheetProtectionObject.unprotect();
```

#### Parameters
None

#### Returns
void
