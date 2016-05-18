# RangeViewCollection Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents a collection of all the RangeView objects.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[RangeView[]](rangeview.md)|A collection of RangeView objects. Read-only.|1.3||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getItemAt(index: number)](#getitematindex-number)|[RangeView](rangeview.md)|Gets a rangeView of a certain row based on its position in the collection.|1.3|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.3|

## Method Details


### getItemAt(index: number)
Gets a rangeView of a certain row based on its position in the collection.

#### Syntax
```js
rangeViewCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[RangeView](rangeview.md)

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

## Example
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var column = table.columns.getItemAt(0);
    column.filter.applyTopItemsFilter(10);
    return ctx.sync()
          .then(function () {
               var visibleValues = table.getDataBodyRange().visibleView;
               visibleValues.load("values");
               return visibleValues;
          })
          .then(ctx.sync)
          .then(function (visibleValues) {
               console.log(JSON.stringify(visibleValues.values));
          });

}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});


// Note:  The visibleValues could also be passed through via:
    .then(function () {
        var visibleValues = table.getDataBodyRange().visibleView;
        visibleValues.load("values");
        ctx.sync(visibleValues);
    })

//===============================================================================

// Single row:

Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var column = table.columns.getItemAt(0);
    column.filter.applyTopItemsFilter(10);
    return ctx.sync()
          .then(function () {
                var firstVisibleRow = table.getDataBodyRange.visibleView.rows.getItemAt(0);
                firstVisibleRow.load("values");
                return firstVisibleRow;
          })
          .then(ctx.sync)
          .then(function (firstVisibleRow) {
               console.log(JSON.stringify(firstVisibleRow.values));
          });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});


// Multiple rows (top 3)

Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var column = table.columns.getItemAt(0);
    column.filter.applyTopItemsFilter(10);
    return ctx.sync()
          .then(function () {
                var visibleTop3Rows = table.getDataBodyRange().visibleView.rows;
                visibleTop3Rows.load({select: "values", top: 3});
                return visibleTop3Rows;
          })
          .then(ctx.sync)
          .then(function (visibleTop3Rows) {
                for (var i = 0; i < visibleTop3Rows.items.length; i++) {
                    console.log(JSON.stringify(visibleTop3Rows.items[0].values));
                }
          });
          
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});

```
