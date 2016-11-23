# TableCollection object (JavaScript API for Excel)

Represents a collection of all the tables that are part of the workbook.

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
|[add(address: Range or string, hasHeaders: bool)](#addaddress-range-or-string-hasheaders-bool)|[Table](table.md)|Create a new table. The range object or source address determines the worksheet under which the table will be added. If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[Table](table.md)|Gets a table by Name or ID.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Table](table.md)|Gets a table based on its position in the collection.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(key: number or string)](#getitemornullkey-number-or-string)|[Table](table.md)|Gets a table by Name or ID. If the table does not exist, the return object's isNull property will be true.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(address: Range or string, hasHeaders: bool)
Create a new table. The range object or source address determines the worksheet under which the table will be added. If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown.

#### Syntax
```js
tableCollectionObject.add(address, hasHeaders);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|address|Range or string|A Range object, or a string address or name of the range representing the data source. If the address does not contain a sheet name, the currently-active sheet is used. Need requirement set 1.1 for string parameter; 1.3 for accepting a Range object.|
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

### getItem(key: number or string)
Gets a table by Name or ID.

#### Syntax
```js
tableCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
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
|:---------------|:--------|:----------|:---|
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


### getItemOrNull(key: number or string)
Gets a table by Name or ID. If the table does not exist, the return object's isNull property will be true.

#### Syntax
```js
tableCollectionObject.getItemOrNull(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|number or string|Name or ID of the table to be retrieved.|

#### Returns
[Table](table.md)

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