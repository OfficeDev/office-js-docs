# WorksheetProtection Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

Represents the protection of a sheet object.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|protected|bool|Indicates if the worksheet is protected. Read-Only. Read-only.|

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

#### Examples
```js
Excel.run(function (ctx) { 
	var sheet = ctx.workbook.worksheets.getItem("Sheet1");
	var range = sheet.getRange("A1:B3").format.protection.locked = false;
	sheet.protection.protect({allowInsertRows:true});
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});

```
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
