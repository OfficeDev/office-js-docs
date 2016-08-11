# ConditionalFormatCollection Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents a collection of all the conditional formats that are active on the current range.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[ConditionalFormat[]](conditionalformat.md)|A collection of conditionalFormat objects. Read-only.|1.3||

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(type: string)](#addtype-string)|[ConditionalFormat](conditionalformat.md)|Adds a new conditional format to the collection, at first priority.|1.3|
|[clearAll()](#clearall)|void|Clears all conditional formats active on the current specified range.|1.3|
|[getCount()](#getcount)|int|Gets the number of items within this conditional format collection.|1.3|
|[getItemAt(index: number)](#getitematindex-number)|[ConditionalFormat](conditionalformat.md)|Gets a conditional formt object based on its position in the collection.|1.3|
|[getItemAtOrNull(index: number)](#getitematornullindex-number)|[ConditionalFormat](conditionalformat.md)|Gets a conditional formt object based on its position in the collection.|1.3|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### add(type: string)
Adds a new conditional format to the collection, at first priority.

#### Syntax
```js
conditionalFormatCollectionObject.add(type);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|type|string|The type of conditional format being added.  Possible values are: Custom, DataBar, ColorScale, IconSet|

#### Returns
[ConditionalFormat](conditionalformat.md)

#### Examples
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    conditionalFormat.iconOrNull.style = "YellowThreeArrows";
    return ctx.sync().then(function () {
        console.log("Added new yellow three arrow icon set.");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```


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
Gets the number of items within this conditional format collection.

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
Gets a conditional formt object based on its position in the collection.

#### Syntax
```js
conditionalFormatCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

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


### getItemAtOrNull(index: number)
Gets a conditional formt object based on its position in the collection.

#### Syntax
```js
conditionalFormatCollectionObject.getItemAtOrNull(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[ConditionalFormat](conditionalformat.md)

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
