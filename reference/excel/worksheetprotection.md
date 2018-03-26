# WorksheetProtection Object (JavaScript API for Excel)

Represents the protection of a sheet object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|protected|bool|Indicates if the worksheet is protected. Read-Only. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|options|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|Sheet protection options. Read-Only. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[protect(options: WorksheetProtectionOptions, password: string)](#protectoptions-worksheetprotectionoptions-password-string)|void|Protects a worksheet. Fails if the worksheet has already been protected.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[unprotect(password: string)](#unprotectpassword-string)|void|Unprotects a worksheet.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### protect(options: WorksheetProtectionOptions, password: string)
Protects a worksheet. Fails if the worksheet has already been protected.

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
Unprotects a worksheet.

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
