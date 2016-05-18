# TableRowCollection Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents a collection of all the rows that are part of the table.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|count|int|Returns the number of rows in the table. Read-only.|1.1||
|items|[TableRow[]](tablerow.md)|A collection of tableRow objects. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(index: number, values: (boolean or string or number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableRow](tablerow.md)|Adds a new row to the table.|1.1|
|[addEmptyRows(index: number, count: number)](#addemptyrowsindex-number-count-number)|[TableRow](tablerow.md)|Adds multiple empty rows to the table and return the first row that is added.|1.3|
|[addMultipleRows(index: number, values: (boolean or string or number)[][])](#addmultiplerowsindex-number-values-boolean-or-string-or-number)|[TableRow](tablerow.md)|Adds multiple new rows to the table. Return the first row that is added.|1.3|
|[getItemAt(index: number)](#getitematindex-number)|[TableRow](tablerow.md)|Gets a row based on its position in the collection.|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### add(index: number, values: (boolean or string or number)[][])
Adds a new row to the table.

#### Syntax
```js
tableRowCollectionObject.add(index, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|Optional. Specifies the relative position of the new row. If null, the addition happens at the end. Any rows below the inserted row are shifted downwards. Zero-indexed.|
|values|(boolean or string or number)[][]|Optional. A 2-dimensional array of unformatted values of the table row.|

#### Returns
[TableRow](tablerow.md)

### addEmptyRows(index: number, count: number)
Adds multiple empty rows to the table and return the first row that is added.

#### Syntax
```js
tableRowCollectionObject.addEmptyRows(index, count);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|Optional. Specifies the relative position of the new row. If null, the addition happens at the end. Any rows below the inserted row are shifted downwards. Zero-indexed.|
|count|number|Specifies the number of empty rows to be added.|

#### Returns
[TableRow](tablerow.md)


#### Example
```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    var rows = tables.getItem("Table1").rows.addEmptyRows(3);
    row.load('index');
    return ctx.sync().then(function() {
        console.log(row.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### addMultipleRows(index: number, values: (boolean or string or number)[][])
Adds multiple new rows to the table. Return the first row that is added.

#### Syntax
```js
tableRowCollectionObject.addMultipleRows(index, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|Optional. Specifies the relative position of the new row. If null, the addition happens at the end. Any rows below the inserted row are shifted downwards. Zero-indexed.|
|values|(boolean or string or number)[][]|A 2-dimensional array of unformatted values of the table row.|

#### Returns
[TableRow](tablerow.md)

#### Example
```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    var rows = tables.getItem("Table1").rows.addMultipleRows(-1,[["A","B"],["C","D"]]);
    row.load('index');
    return ctx.sync().then(function() {
        console.log(row.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### getItemAt(index: number)
Gets a row based on its position in the collection.

#### Syntax
```js
tableRowCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[TableRow](tablerow.md)

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
