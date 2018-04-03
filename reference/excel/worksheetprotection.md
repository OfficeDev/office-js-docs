# WorksheetProtection Object (JavaScript API for Excel)

Represents the protection of a sheet object.

## Properties

| Property  | Type | Description                                                    | Req. Set                                                 |
| :-------- | :--- | :------------------------------------------------------------- | :------------------------------------------------------- |
| protected | bool | Indicates if the worksheet is protected. Read-Only. Read-only. | [1.2](../requirement-sets/excel-api-requirement-sets.md) |

## Relationships

| Relationship | Type                           | Description                                     | Req. Set                                                 |
| :----------- | :----------------------------- | :---------------------------------------------- | :------------------------------------------------------- |
| options      | [WorksheetProtectionOptions][] | Sheet protection options. Read-Only. Read-only. | [1.2](../requirement-sets/excel-api-requirement-sets.md) |

## Methods

| Method                                                                                                                       | Return Type | Description                                                              | Req. Set                                                 |
| :--------------------------------------------------------------------------------------------------------------------------- | :---------- | :----------------------------------------------------------------------- | :------------------------------------------------------- |
| [protect(options: WorksheetProtectionOptions, password: string)](#protectoptions-worksheetprotectionoptions-password-string) | void        | Protects a worksheet. Fails if the worksheet has already been protected. | [1.2](../requirement-sets/excel-api-requirement-sets.md) |
| [unprotect(password: string)](#unprotectpassword-string)                                                                     | void        | Unprotects a worksheet.                                                  | [1.2](../requirement-sets/excel-api-requirement-sets.md) |

## Method Details

### protect(options: WorksheetProtectionOptions, password: string)

Protects a worksheet. Fails if the worksheet has already been protected.

#### Syntax

```js
worksheetProtectionObject.protect(options, password);
```

#### Parameters

| Parameter | Type                       | Description                          |
| :-------- | :------------------------- | :----------------------------------- |
| options   | WorksheetProtectionOptions | Optional. sheet protection options.  |
| password  | string                     | Optional. sheet protection password. |

#### Returns

void

#### Examples

```js
Excel.run(function(ctx) {
  // get a reference to Sheet1
  var sheet = ctx.workbook.worksheets.getItem("Sheet1");

  // Protect inserting or deleting rows in Sheet1
  sheet.protection.protect({
    allowInsertRows: false,
    allowDeleteRows: false
  });

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

| Parameter | Type   | Description                          |
| :-------- | :----- | :----------------------------------- |
| password  | string | Optional. sheet protection password. |

#### Returns

void

#### Examples

```js
Excel.run(function(ctx) {
  // get a reference to Sheet1
  var sheet = ctx.workbook.worksheets.getItem("Sheet1");

  // Remove all protects applied to Sheet1
  sheet.protection.unprotect();

  return ctx.sync();
}).catch(function(error) {
  console.log("Error: " + error);
  if (error instanceof OfficeExtension.Error) {
    console.log("Debug info: " + JSON.stringify(error.debugInfo));
  }
});
```

#### Remakes

Unprotecting a worksheet with `unprotect()` will remove _all_ [WorksheetProtectionOptions][] options applied to a worksheet.
To remove only a subset of WorksheetProtectionOptions use the `protect()` method and set the options you wish to remove to `true`.

```js
Excel.run(function(ctx) {
  var sheet = ctx.workbook.worksheets.getItem("Sheet1");
  sheet.protection.protect({
    allowInsertRows: false, // Protect row insertion
    allowDeleteRows: true // Unprotect row deletion
  });
});
```

[worksheetprotectionoptions]: worksheetprotectionoptions.md
