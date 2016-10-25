# TableCell Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Represents a table cell in a Word document.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|cellIndex|int|Gets the index of the cell in its row. Read-only.|[1.3](../reqset/word-requirement.md)|
|columnWidth|float|Gets and sets the width of the cell's column in points. This is applicable to uniform tables.|[1.3](../reqset/word-requirement.md)|
|rowIndex|int|Gets the index of the cell's row in the table. Read-only.|[1.3](../reqset/word-requirement.md)|
|shadingColor|string|Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.|[1.3](../reqset/word-requirement.md)|
|value|string|Gets and sets the text of the cell.|[1.3](../reqset/word-requirement.md)|
|width|float|Gets the width of the cell in points. Read-only.|[1.3](../reqset/word-requirement.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|body|[Body](body.md)|Gets the body object of the cell. Read-only.|[1.3](../reqset/word-requirement.md)|
|horizontalAlignment|[Alignment](alignment.md)|Gets and sets the horizontal alignment of the cell. The value can be 'left', 'centered', 'right', or 'justified'.|[1.3](../reqset/word-requirement.md)|
|parentRow|[TableRow](tablerow.md)|Gets the parent row of the cell. Read-only.|[1.3](../reqset/word-requirement.md)|
|parentTable|[Table](table.md)|Gets the parent table of the cell. Read-only.|[1.3](../reqset/word-requirement.md)|
|verticalAlignment|[VerticalAlignment](verticalalignment.md)|Gets and sets the vertical alignment of the cell. The value can be 'top', 'center' or 'bottom'.|[1.3](../reqset/word-requirement.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[deleteColumn()](#deletecolumn)|void|Deletes the column containing this cell. This is applicable to uniform tables.|[1.3](../reqset/word-requirement.md)|
|[deleteRow()](#deleterow)|void|Deletes the row containing this cell.|[1.3](../reqset/word-requirement.md)|
|[getBorder(borderLocation: BorderLocation)](#getborderborderlocation-borderlocation)|[TableBorder](tableborder.md)|Gets the border style for the specified border.|[1.3](../reqset/word-requirement.md)|
|[getCellPadding(cellPaddingLocation: CellPaddingLocation)](#getcellpaddingcellpaddinglocation-cellpaddinglocation)|[float?](float?.md)|Gets cell padding in points.|[1.3](../reqset/word-requirement.md)|
|[getNext()](#getnext)|[TableCell](tablecell.md)|Gets the next cell.|[1.3](../reqset/word-requirement.md)|
|[insertColumns(insertLocation: InsertLocation, columnCount: number, values: string[][])](#insertcolumnsinsertlocation-insertlocation-columncount-number-values-string)|void|Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.|[1.3](../reqset/word-requirement.md)|
|[insertRows(insertLocation: InsertLocation, rowCount: number, values: string[][])](#insertrowsinsertlocation-insertlocation-rowcount-number-values-string)|[TableRowCollection](tablerowcollection.md)|Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.|[1.3](../reqset/word-requirement.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../reqset/word-requirement.md)|
|[setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)](#setcellpaddingcellpaddinglocation-cellpaddinglocation-cellpadding-float)|void|Sets cell padding in points.|[1.3](../reqset/word-requirement.md)|
|[split(rowCount: number, columnCount: number)](#splitrowcount-number-columncount-number)|void|Adds columns to the left or right of the cell, using the existing column as a template. The string values, if specified, are set in the newly inserted rows.|[1.4](../reqset/word-requirement.md)|

## Method Details


### deleteColumn()
Deletes the column containing this cell. This is applicable to uniform tables.

#### Syntax
```js
tableCellObject.deleteColumn();
```

#### Parameters
None

#### Returns
void

### deleteRow()
Deletes the row containing this cell.

#### Syntax
```js
tableCellObject.deleteRow();
```

#### Parameters
None

#### Returns
void

### getBorder(borderLocation: BorderLocation)
Gets the border style for the specified border.

#### Syntax
```js
tableCellObject.getBorder(borderLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|borderLocation|BorderLocation|Required. The border location.|

#### Returns
[TableBorder](tableborder.md)

### getCellPadding(cellPaddingLocation: CellPaddingLocation)
Gets cell padding in points.

#### Syntax
```js
tableCellObject.getCellPadding(cellPaddingLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|cellPaddingLocation|CellPaddingLocation|Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.|

#### Returns
[float?](float?.md)

### getNext()
Gets the next cell.

#### Syntax
```js
tableCellObject.getNext();
```

#### Parameters
None

#### Returns
[TableCell](tablecell.md)

### insertColumns(insertLocation: InsertLocation, columnCount: number, values: string[][])
Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.

#### Syntax
```js
tableCellObject.insertColumns(insertLocation, columnCount, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|insertLocation|InsertLocation|Required. It can be 'Before' or 'After'.|
|columnCount|number|Required. Number of columns to add|
|values|string[][]|Optional. Optional 2D array. Cells are filled if the corresponding strings are specified in the array.|

#### Returns
void

### insertRows(insertLocation: InsertLocation, rowCount: number, values: string[][])
Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.

#### Syntax
```js
tableCellObject.insertRows(insertLocation, rowCount, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|insertLocation|InsertLocation|Required. It can be 'Before' or 'After'.|
|rowCount|number|Required. Number of rows to add.|
|values|string[][]|Optional. Optional 2D array. Cells are filled if the corresponding strings are specified in the array.|

#### Returns
[TableRowCollection](tablerowcollection.md)

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

### setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)
Sets cell padding in points.

#### Syntax
```js
tableCellObject.setCellPadding(cellPaddingLocation, cellPadding);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|cellPaddingLocation|CellPaddingLocation|Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.|
|cellPadding|float|Required. The cell padding.|

#### Returns
void

### split(rowCount: number, columnCount: number)
Adds columns to the left or right of the cell, using the existing column as a template. The string values, if specified, are set in the newly inserted rows.

#### Syntax
```js
tableCellObject.split(rowCount, columnCount);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|rowCount|number|Required. The number of rows to split into. Must be a divisor of the number of underlying rows.|
|columnCount|number|Required. The number of columns to split into.|

#### Returns
void
