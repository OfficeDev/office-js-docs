# TableRow Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  

Represents a row in a table.

To provide feedback on this API, you can [file an issue in GitHub](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow).

## Properties

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|cellCount|int|Gets the number of cells in the row. Read-only.|
|id|string|Gets the ID of the row. Read-only.|
|rowIndex|int|Gets the index of the row in its parent table. Read-only.|

_See [property access examples](#property-access-examples)_.

## Relationships

| Relationship | Type	|Description| 
|:---------------|:--------|:----------|
|cells|[TableCellCollection](tablecellcollection.md)|Gets the cells in the row. Read-only.|
|parentTable|[Table](table.md)|Gets the parent table. Read-only.|

## Methods

| Method		   | Return Type	|Description| 
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Clears the contents of the row.|
|[insertRowAsSibling(insertLocation: string, values: string[])](#insertrowassiblinginsertlocation-string-values-string)|[TableRow](tablerow.md)|Inserts a row before or after the current row.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.|
|[setShadingColor(colorCode: string)](#setshadingcolorcolorcode-string)|void|Sets the shading color of all cells in the row.|

## Method details

### clear()

Clears the contents of the row.

#### Syntax

```js
tableRowObject.clear();
```

#### Parameters

None

#### Returns

Void

<br/>

### insertRowAsSibling(insertLocation: string, values: string[])

Inserts a row before or after the current row.

#### Syntax

```js
tableRowObject.insertRowAsSibling(insertLocation, values);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|insertLocation|string|Where the new rows should be inserted relative to the current row. Possible values are Before, After.|
|values|string[]|Optional. Strings to insert in the new row, specified as an array. Must not have more cells than in the current row.|

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
		
		// for each table, get table rows.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				
				// Queue a command to load table.rows.
				ctx.load(table, "rows");
				
				// Run the queued commands
				return ctx.sync().then(function() {
					var rows = table.rows;
					rows.items[1].insertRowAsSibling("Before", ["cell0", "cell1"]);
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

Sets the shading color of all cells in the row.

#### Syntax

```js
tableRowObject.setShadingColor(colorCode);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|colorCode|string|The color code to set the cells to.|

#### Returns

Void

<br/>

### Property access examples

**id, cellCount, rowIndex**

```js
OneNote.run(function(ctx) {
	var app = ctx.application;
	var outline = app.getActiveOutline();
	
	// Queue a command to load outline.paragraphs and their types.
	ctx.load(outline, "paragraphs, paragraphs/type");
	
	// Run the queued commands, and return a promise to indicate task completion.
	return ctx.sync().then(function () {
		var paragraphs = outline.paragraphs;
		
		// for each table, get table rows.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				
				// Queue a command to load table.rows.
				ctx.load(table, "rows");
				return ctx.sync().then(function() {
					var rows = table.rows;
					
					// for each table row, log cell count and row index.
					for (var i = 0; i < rows.items.length; i++) {
						console.log("Row " + i + " Id: " + rows.items[i].id);
						console.log("Row " + i + " Cell Count: " + rows.items[i].cellCount);
						console.log("Row " + i + " Row Index: " + rows.items[i].rowIndex);
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

**parentTable, cells**

```js
OneNote.run(function(ctx) {
	var app = ctx.application;
	var outline = app.getActiveOutline();
	
	// Queue a command to load outline.paragraphs and their types.
	ctx.load(outline, "paragraphs, paragraphs/type");
	
	// Run the queued commands, and return a promise to indicate task completion.
	return ctx.sync().then(function () {
		var paragraphs = outline.paragraphs;
		
		// for each table, get table rows.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				
				// Queue a command to load parentTable and cells of each row in the table.
				ctx.load(table, "rows/parentTable, rows/cells");
				return ctx.sync().then(function() {
					var rows = table.rows;
					
					// for each row, log parentTable and cells
					for (var i = 0; i < rows.items.length; i++) {
						console.log("Row " + i + " Parent Table Id: " + rows.items[i].parentTable.id);
						var cells = rows.items[i].cells;
						for (var j = 0 ; j < cells.items.length; j++) {
							console.log("Row " + i + " Cell " + j + " Id: " + cells.items[j].id);
						}
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

