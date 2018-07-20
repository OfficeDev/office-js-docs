# WorkbookProtection Object (JavaScript API for Excel)

Represents the protection of a workbook object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|protected|bool|Indicates if the workbook is protected. Read-Only. Read-only.|[ApiSet.InProgressFeatures.TemporaryMovedTo18](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[protect(password: string)](#protectpassword-string)|void|Protects a workbook. Fails if the workbook has been protected.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[unprotect(password: string)](#unprotectpassword-string)|void|Unprotects a workbook.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### protect(password: string)
Protects a workbook. Fails if the workbook has been protected.

#### Syntax
```js
workbookProtectionObject.protect(password);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|password|string|Optional. workbook protection password.|

#### Returns
void

### unprotect(password: string)
Unprotects a workbook.

#### Syntax
```js
workbookProtectionObject.unprotect(password);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|password|string|Optional. workbook protection password.|

#### Returns
void
