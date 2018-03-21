# WorkbookCreated Object (JavaScript API for Excel)

The WorkbookCreated object is the top level object created by Application.CreateWorkbook. A WorkbookCreated object is a special Workbook object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|Returns a value that uniquely identifies the WorkbookCreated object. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[open()](#open)|void|Open the workbook.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### open()
Open the workbook.

#### Syntax
```js
workbookCreatedObject.open();
```

#### Parameters
None

#### Returns
void
