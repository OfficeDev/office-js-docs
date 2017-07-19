# TableCollection Object (JavaScript API for Excel)

Represents a collection of all the tables that are part of the workbook or worksheet, depending on how it was reached.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|count|int|Returns the number of tables in the workbook. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[Table[]](table.md)|A collection of table objects. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(address: object, hasHeaders: bool)](#addaddress-object-hasheaders-bool)|[Table](table.md)|Create a new table. The range object or source address determines the worksheet under which the table will be added. If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Gets the number of tables in the collection.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[Table](table.md)|Gets a table by Name or ID.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Table](table.md)|Gets a table based on its position in the collection.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: number or string)](#getitemornullobjectkey-number-or-string)|[Table](table.md)|Gets a table by Name or ID. If the table does not exist, will return a null object.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|


## Events

| Event		 | Argument Type of Event handler |Description| Req. Set|
|:---------------|:----------|:----|:----|
|[onDataChanged](#ondatachanged)|[TableDataChangedEvent](tabledatachangedevent.md)|Occurs when data changed on any table in a workbook, or a worksheet.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(address: object, hasHeaders: bool)
Create a new table. The range object or source address determines the worksheet under which the table will be added. If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown.

#### Syntax
```js
tableCollectionObject.add(address, hasHeaders);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|address|object|A Range object, or a string address or name of the range representing the data source. If the address does not contain a sheet name, the currently-active sheet is used.|
|hasHeaders|bool|Boolean value that indicates whether the data being imported has column labels. If the source does not contain headers (i.e,. when this property set to false), Excel will automatically generate header shifting the data down by one row.|

#### Returns
[Table](table.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.add('Sheet1!A1:E7', true);
	table.load('name');
	return ctx.sync().then(function() {
		console.log(table.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getCount()
Gets the number of tables in the collection.

#### Syntax
```js
tableCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(key: number or string)
Gets a table by Name or ID.

#### Syntax
```js
tableCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|number or string|Name or ID of the table to be retrieved.|

#### Returns
[Table](table.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	table.load('name');
	return ctx.sync().then(function() {
			console.log(table.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


#### Examples

```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.getItemAt(0);
	table.load('name');
	return ctx.sync().then(function() {
			console.log(table.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getItemAt(index: number)
Gets a table based on its position in the collection.

#### Syntax
```js
tableCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[Table](table.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.getItemAt(0);
	table.load('name');
	return ctx.sync().then(function() {
			console.log(table.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getItemOrNullObject(key: number or string)
Gets a table by Name or ID. If the table does not exist, will return a null object.

#### Syntax
```js
tableCollectionObject.getItemOrNullObject(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|number or string|Name or ID of the table to be retrieved.|

#### Returns
[Table](table.md)
### Property access examples

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	tables.load();
	return ctx.sync().then(function() {
		console.log("tables Count: " + tables.count);
		for (var i = 0; i < tables.items.length; i++)
		{
			console.log(tables.items[i].name);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Get the number of tables

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	tables.load('count');
	return ctx.sync().then(function() {
		console.log(tables.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

## Event Details

### onDataChanged
Occurs when data changed on any table in a workbook, or a worksheet.

#### Syntax
```js
tables.onDataChanged.add(onTablesDataChanged);
```
#### Example
```js
Excel.run(function (ctx) {
    ctx.workbook.tables.onDataChanged.add(onTablesDataChanged);

    return ctx.sync().then(function () {
        console.log("add event succeed.");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});

function onTablesDataChanged(event) {
    return Excel.run(event, function (context) {
        var table = event.table;
        var range = event.range;
        var source = event.source;
        var changeType = event.changeType;
        var columns = table.columns;
        var rows = talbe.rows;
        var worksheet = table.worksheet;

        //Load table related properties
        table.load("name");
        columns.load("count");
        rows.load("count");
        worksheet.load("name");
        range.load("address, cellCount");

        return context.sync().then(function () {
            // Get the table that has data change
            var columnCount = columns.count;
            var rowCount = rows.count;
            console.log("Data Changed in worksheet " + worksheet.name + "Table name is " +                table.name + "Total column number of table is" + columnCount + "Total row number of table is" + rowCount);
            // Get the changed range
            console.log(".Changed area address is " + range.address +
                "cell count in the range is" + range.cellCount);
            //Get the event source, data change comes from local or remote.
            console.log("Data Change comes from" + source);
            
            //get the new inserted Row.
            else if (changeType == "RowInsert") {
                console.log("Row inserted on row" + range.address);
            }
            //else if ... //For RowDelete, ColumnInsert, ColumnDelete, same as above.

            // Other changes happened
            else {
                console.log("Other changes happened");
            }
        });
    });
}

```
