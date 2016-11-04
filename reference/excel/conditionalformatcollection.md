# ConditionalFormatCollection Object (JavaScript API for Excel)

Represents a collection of all the conditional formats that are overlap the range.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[ConditionalFormat[]](conditionalformat.md)|A collection of conditionalFormat objects. Read-only.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clearAll()](#clearall)|void|Clears all conditional formats active on the current specified range.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Returns the number of conditional formats in the workbook. Read-only.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[ConditionalFormat](conditionalformat.md)|Returns a conditional format at the given index.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### clearAll()
Clears all conditional formats active on the current specified range.

#### Syntax
```js
conditionalFormatCollectionObject.clearAll();
```

#### Parameters
None

#### Returns
void

#### Examples
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormats = range.conditionalFormats;
    var conditionalFormat = conditionalFormats.clearAll();
    return ctx.sync().then(function () {
        console.log("Cleared all conditional formats from this range.");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```


### getCount()
Returns the number of conditional formats in the workbook. Read-only.

#### Syntax
```js
conditionalFormatCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

#### Examples
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    conditionalFormat.iconOrNull.style = Excel.IconSet.fourTrafficLights;
    var cfCount = range.conditionalFormats.getCount(); 

    return ctx.sync().then(function () {
        console.log("Count: " + cfCount.value);
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### getItemAt(index: number)
Returns a conditional format at the given index.

#### Syntax
```js
conditionalFormatCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|Index of the conditional formats to be retrieved.|

#### Returns
[ConditionalFormat](conditionalformat.md)

#### Examples
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormats = range.conditionalFormats;
    var conditionalFormat = conditionalFormats.getItemAt(3);
    return ctx.sync().then(function () {
        console.log("Conditional Format at Item 3 Loaded");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```


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
