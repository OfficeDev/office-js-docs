# Table Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  

Represents a table in a OneNote page.

To provide feedback on this API, you can [file an issue in GitHub](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table).

## Properties

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|borderVisible|bool|Gets or sets whether the borders are visible or not. True if they are visible, false if they are hidden.|
|columnCount|int|Gets the number of columns in the table. Read-only.|
|id|string|Gets the ID of the table. Read-only.|
|rowCount|int|Gets the number of rows in the table. Read-only.|

_See [property access examples](#property-access-examples)_.

## Relationships

| Relationship | Type	|Description| 
|:---------------|:--------|:----------|
|paragraph|[Paragraph](paragraph.md)|Gets the Paragraph object that contains the Table object. Read-only.|
|rows|[TableRowCollection](tablerowcollection.md)|Gets all the table rows. Read-only.|

## Methods

| Method		   | Return Type	|Description| 
|:---------------|:--------|:----------|
|[appendColumn(values: string[])](#appendcolumnvalues-string)|void|Adds a column to the end of the table. Values, if specified, are set in the new column. Otherwise, the column is empty.|
|[appendRow(values: string[])](#appendrowvalues-string)|[TableRow](tablerow.md)|Adds a row to the end of the table. Values, if specified, are set in the new row. Otherwise, the row is empty.|
|[clear()](#clear)|void|Clears the contents of the table.|
|[getCell(rowIndex: number, cellIndex: number)](#getcellrowindex-number-cellindex-number)|[TableCell](tablecell.md)|Gets the table cell at a specified row and column.|
|[insertColumn(index: number, values: string[])](#insertcolumnindex-number-values-string)|void|Inserts a column at the given index in the table. Values, if specified, are set in the new column. Otherwise, the column is empty.|
|[insertRow(index: number, values: string[])](#insertrowindex-number-values-string)|[TableRow](tablerow.md)|Inserts a row at the given index in the table. Values, if specified, are set in the new row. Otherwise, the row is empty.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.|
|[setShadingColor(colorCode: string)](#setshadingcolorcolorcode-string)|void|Sets the shading color of all cells in the table.|

## Method details

### appendColumn(values: string[])

Adds a column to the end of the table. Values, if specified, are set in the new column. Otherwise, the column is empty.

#### Syntax

```js
tableObject.appendColumn(values);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|values|string[]|Optional. Strings to insert in the new column, specified as an array. Must not have more values than rows in the table.|

#### Returns

Void

#### Examples

```js
OneNote.run(function(ctx) {
	var app = ctx.application;
	var outline = app.getActiveOutline();
	
	// Queue a command to load outline.paragraphs and their types.
	ctx.load(outline, "paragraphs, paragraphs/type");
	
	// Run the queued commands, and return a promise to indicate task completion.
	return ctx.sync().then(function () {
		var paragraphs = outline.paragraphs;
		
		// for each table, append a column.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				table.appendColumn(["cell0", "cell1"]);
			}
		}
		return ctx.sync();
	})
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>

### appendRow(values: string[])

Adds a row to the end of the table. Values, if specified, are set in the new row. Otherwise, the row is empty.

#### Syntax

```js
tableObject.appendRow(values);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|values|string[]|Optional. Strings to insert in the new row, specified as an array. Must not have more values than columns in the table.|

#### Returns

[TableRow](tablerow.md)

#### Examples

```js
OneNote.run(function(ctx) {
	var app = ctx.application;
	var outline = app.getActiveOutline();
	
	// Queue a command to load outline.paragraphs and their types.
	ctx.load(outline, "paragraphs, paragraphs/type");
	
	// Run the queued commands, and return a promise to indicate task completion.
	return ctx.sync().then(function () {
		var paragraphs = outline.paragraphs;
		
		// for each table, append a column.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				var row = table.appendRow(["cell0", "cell1"]);
			}
		}
		return ctx.sync();
	})
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>

### clear()

Clears the contents of the table.

#### Syntax

```js
tableObject.clear();
```

#### Parameters

None

#### Returns

Void

<br/>

### getCell(rowIndex: number, cellIndex: number)

Gets the table cell at a specified row and column.

#### Syntax

```js
tableObject.getCell(rowIndex, cellIndex);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rowIndex|number|The index of the row.|
|cellIndex|number|The index of the cell in the row.|

#### Returns

[TableCell](tablecell.md)


#### Examples

```js
OneNote.run(function(ctx) {
	var app = ctx.application;
	var outline = app.getActiveOutline();
	
	// Queue a command to load outline.paragraphs and their types.
	ctx.load(outline, "paragraphs, paragraphs/type");
	
	// Run the queued commands, and return a promise to indicate task completion.
	return ctx.sync().then(function () {
		var paragraphs = outline.paragraphs;
		
		// for each table, get a cell in the second row and third column.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				var cell = table.getCell(2 /*Row Index*/, 3 /*Column Index*/);
			}
		}
		return ctx.sync();
	})
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>

### insertColumn(index: number, values: string[])

Inserts a column at the given index in the table. Values, if specified, are set in the new column. Otherwise, the column is empty.

#### Syntax

```js
tableObject.insertColumn(index, values);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index where the column will be inserted in the table.|
|values|string[]|Optional. Strings to insert in the new column, specified as an array. Must not have more values than rows in the table.|

#### Returns

Void

#### Examples

```js
OneNote.run(function(ctx) {
	var app = ctx.application;
	var outline = app.getActiveOutline();
	
	// Queue a command to load outline.paragraphs and their types.
	ctx.load(outline, "paragraphs, paragraphs/type");
	
	// Run the queued commands, and return a promise to indicate task completion.
	return ctx.sync().then(function () {
		var paragraphs = outline.paragraphs;
		
		// for each table, insert a column at index two.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				table.insertColumn(2, ["cell0", "cell1"]);
			}
		}
		return ctx.sync();
	})
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>

### insertRow(index: number, values: string[])

Inserts a row at the given index in the table. Values, if specified, are set in the new row. Otherwise, the row is empty.

#### Syntax

```js
tableObject.insertRow(index, values);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index where the row will be inserted in the table.|
|values|string[]|Optional. Strings to insert in the new row, specified as an array. Must not have more values than columns in the table.|

#### Returns

[TableRow](tablerow.md)

#### Examples

```js
OneNote.run(function(ctx) {
	var app = ctx.application;
	var outline = app.getActiveOutline();
	
	// Queue a command to load outline.paragraphs and their types.
	ctx.load(outline, "paragraphs, paragraphs/type");
	
	// Run the queued commands, and return a promise to indicate task completion.
	return ctx.sync().then(function () {
		var paragraphs = outline.paragraphs;
		
		// for each table, insert a row at index two.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				var row = table.insertRow(2, ["cell0", "cell1"]);
			}
		}
		return ctx.sync();
	})
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>

### load(param: object)

Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.

#### Syntax

```js
object.load(param);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as a delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns

Void

<br/>

### setShadingColor(colorCode: string)

Sets the shading color of all cells in the table.

#### Syntax

```js
tableObject.setShadingColor(colorCode);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|colorCode|string|The color code to set the cells to.|

#### Returns

Void

<br/>

### Property access examples

**columnCount, rowCount, id**

```js
OneNote.run(function(ctx) {
	var app = ctx.application;
	var outline = app.getActiveOutline();
	
	// Queue a command to load outline.paragraphs and their types.
	ctx.load(outline, "paragraphs/type");
	
	// Run the queued commands, and return a promise to indicate task completion.
	return ctx.sync().then(function () {
		var paragraphs = outline.paragraphs;
		
		// For each table, log properties.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				ctx.load(table);
				return ctx.sync().then(function() {
					console.log("Table Id: " + table.id);
					console.log("Row Count: " + table.rowCount);
					console.log("Column Count: " + table.columnCount);
					return ctx.sync();
				});
			}
		}
	});
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>

**paragraph, rows**

```js
OneNote.run(function(ctx) {
	var app = ctx.application;
	var outline = app.getActiveOutline();
	
	// Queue a command to load outline.paragraphs and their types.
	ctx.load(outline, "paragraphs, paragraphs/type");
	
	// Run the queued commands, and return a promise to indicate task completion.
	return ctx.sync().then(function () {
		var paragraphs = outline.paragraphs;
		
		// for each table, log its paragraph id.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				ctx.load(table, "paragraph/id, rows/id");
				return ctx.sync().then(function() {
					console.log("Paragraph Id: " + table.paragraph.id);
					var rows = table.rows;
					
					// for each rows in the table, log row index and id.
					for (var i = 0; i < rows.items.length; i++) {
						console.log("Row " + i + " Id: " + rows.items[i].id);
					}
					return ctx.sync();
				});
			}
		}
	})
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>

