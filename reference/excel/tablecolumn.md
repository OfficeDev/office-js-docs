# TableColumn Object (JavaScript API for Excel)

Represents a column in a table.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|int|Returns a unique key that identifies the column within the table. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|index|int|Returns the index number of the column within the columns collection of the table. Zero-indexed. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Represents the name of the table column.|[1.1, 1.1 for getting the name; 1.4 for setting it.](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|filter|[Filter](filter.md)|Retrieve the filter applied to the column. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the column from the table.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Gets the range object associated with the data body of the column.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Gets the range object associated with the header row of the column.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getNextColumn()](#getnextcolumn)|[TableColumn](tablecolumn.md)|Gets the table column that follows this one. If there are no table columns following this one, this method will throw an error.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getNextColumnOrNullObject()](#getnextcolumnornullobject)|[TableColumn](tablecolumn.md)|Gets the table column that follows this one. If there are no table columns following this one, this method will return a null object.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getPreviousColumn()](#getpreviouscolumn)|[TableColumn](tablecolumn.md)|Gets the table column that precedes this one. If there are no previous table columns, this method will throw an error.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getPreviousColumnOrNullObject()](#getpreviouscolumnornullobject)|[TableColumn](tablecolumn.md)|Gets the table column that precedes this one. If there are no previous table columns, this method will return a null objet.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|Gets the range object associated with the entire column.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Gets the range object associated with the totals row of the column.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### delete()
Deletes the column from the table.

#### Syntax
```js
tableColumnObject.delete();
```

#### Parameters
None

#### Returns
void

#### Examples

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var column = ctx.workbook.tables.getItem(tableName).columns.getItemAt(2);
	column.delete();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getDataBodyRange()
Gets the range object associated with the data body of the column.

#### Syntax
```js
tableColumnObject.getDataBodyRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var column = ctx.workbook.tables.getItem(tableName).columns.getItemAt(0);
	var dataBodyRange = column.getDataBodyRange();
	dataBodyRange.load('address');
	return ctx.sync().then(function() {
		console.log(dataBodyRange.address);
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getHeaderRowRange()
Gets the range object associated with the header row of the column.

#### Syntax
```js
tableColumnObject.getHeaderRowRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var columns = ctx.workbook.tables.getItem(tableName).columns.getItemAt(0);
	var headerRowRange = columns.getHeaderRowRange();
	headerRowRange.load('address');
	return ctx.sync().then(function() {
		console.log(headerRowRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getNextColumn()
Gets the table column that follows this one. If there are no table columns following this one, this method will throw an error.

#### Syntax
```js
tableColumnObject.getNextColumn();
```

#### Parameters
None

#### Returns
[TableColumn](tablecolumn.md)

### getNextColumnOrNullObject()
Gets the table column that follows this one. If there are no table columns following this one, this method will return a null object.

#### Syntax
```js
tableColumnObject.getNextColumnOrNullObject();
```

#### Parameters
None

#### Returns
[TableColumn](tablecolumn.md)

### getPreviousColumn()
Gets the table column that precedes this one. If there are no previous table columns, this method will throw an error.

#### Syntax
```js
tableColumnObject.getPreviousColumn();
```

#### Parameters
None

#### Returns
[TableColumn](tablecolumn.md)

### getPreviousColumnOrNullObject()
Gets the table column that precedes this one. If there are no previous table columns, this method will return a null objet.

#### Syntax
```js
tableColumnObject.getPreviousColumnOrNullObject();
```

#### Parameters
None

#### Returns
[TableColumn](tablecolumn.md)

### getRange()
Gets the range object associated with the entire column.

#### Syntax
```js
tableColumnObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var columns = ctx.workbook.tables.getItem(tableName).columns.getItemAt(0);
	var columnRange = columns.getRange();
	columnRange.load('address');
	return ctx.sync().then(function() {
		console.log(columnRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getTotalRowRange()
Gets the range object associated with the totals row of the column.

#### Syntax
```js
tableColumnObject.getTotalRowRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var columns = ctx.workbook.tables.getItem(tableName).columns.getItemAt(0);
	var totalRowRange = columns.getTotalRowRange();
	totalRowRange.load('address');
	return ctx.sync().then(function() {
		console.log(totalRowRange.address);
	});
}).catch(function(error) {
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
### Property access examples

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var column = ctx.workbook.tables.getItem(tableName).columns.getItem(0);
	column.load('index');
	return ctx.sync().then(function() {
		console.log(column.index);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var tables = ctx.workbook.tables;
	var newValues = [["New"], ["Values"], ["For"], ["New"], ["Column"]];
	var column = ctx.workbook.tables.getItem(tableName).columns.getItemAt(2);
	column.values = newValues;
	column.load('values');
	return ctx.sync().then(function() {
		console.log(column.values);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```