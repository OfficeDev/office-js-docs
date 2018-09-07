# Application Object (JavaScript API for Excel)

Represents the Excel application that manages the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|calculationEngineVersion|int|Returns a number about the version of Excel Calculation Engine that the workbook was last fully recalculated by. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|calculationMode|string|Returns the calculation mode used in the workbook, as defined by the constants in Excel.CalculationMode. Possible values are: `Automatic`, where Excel controls recalculation; `AutomaticExceptTables`, where Excel controls recalculation but ignores changes in tables; `Manual`, where calculation is done when the user requests it. Possible values are: `Automatic` Excel controls recalculation,`AutomaticExceptTables` Excel controls recalculation but ignores changes in tables.,`Manual` Calculation is done when the user requests it.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|calculationState|string|Returns a CalculationState that indicates the calculation state of the application. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|iterativeCalculation|[IterativeCalculation](iterativecalculation.md)|Returns the Iterative Calculation settings. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|Recalculate all currently opened workbooks in Excel.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[createWorkbook(base64File: string)](#createworkbookbase64file-string)|[WorkbookCreated](workbookcreated.md)|Creates a new hidden workbook by using an optional base64 encoded .xlsx file.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[suspendApiCalculationUntilNextSync()](#suspendapicalculationuntilnextsync)|void|Suspends calculation until the next "context.sync()" is called. Once set, it is the developer's responsibility to re-calc the workbook, to ensure that any dependencies are propagated.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|[suspendScreenUpdatingUntilNextSync()](#suspendscreenupdatinguntilnextsync)|void|Suspends sceen updating until the next "context.sync()" is called.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

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
|calculationType|string|Specifies the calculation type to use. Possible values are: `Recalculate` Default-option. Performs normal calculation by calculating all the formulas in the workbook,`Full` Forces a full calculation of the data,`FullRebuild`  Forces a full calculation of the data and rebuilds the dependencies.|

#### Returns
void

#### Examples
```js
Excel.run(function (ctx) {
	ctx.workbook.application.calculate('Full');
	return ctx.sync();
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### createWorkbook(base64File: string)
Creates a new hidden workbook by using an optional base64 encoded .xlsx file.

#### Syntax
```js
applicationObject.createWorkbook(base64File);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64File|string|Optional. Optional. The base64 encoded .xlsx file. The default value is null.|

#### Returns
[WorkbookCreated](workbookcreated.md)

### suspendApiCalculationUntilNextSync()
Suspends calculation until the next "context.sync()" is called. Once set, it is the developer's responsibility to re-calc the workbook, to ensure that any dependencies are propagated.

#### Syntax
```js
applicationObject.suspendApiCalculationUntilNextSync();
```

#### Parameters
None

#### Returns
void

### suspendScreenUpdatingUntilNextSync()
Suspends sceen updating until the next "context.sync()" is called.

#### Syntax
```js
applicationObject.suspendScreenUpdatingUntilNextSync();
```

#### Parameters
None

#### Returns
void
### Property access examples
```js
Excel.run(function (ctx) {
	var application = ctx.workbook.application;
	application.load('calculationMode');
	return ctx.sync().then(function() {
		console.log(application.calculationMode);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

