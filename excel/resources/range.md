# Range Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Range represents a set of one or more contiguous cells such as a cell, a row, a column, block of cells, etc.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|address|string|Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. Sheet1!A1:B4). Read-only.|1.1||
|addressLocal|string|Represents range reference for the specified range in the language of the user. Read-only.|1.1||
|cellCount|int|Number of cells in the range. Read-only.|1.1||
|columnCount|int|Represents the total number of columns in the range. Read-only.|1.1||
|columnHidden|bool|Represents if all columns of the current range are hidden.|1.2||
|columnIndex|int|Represents the column number of the first cell in the range. Zero-indexed. Read-only.|1.1||
|formulas|object[][]|Represents the formula in A1-style notation.|1.1||
|formulasLocal|object[][]|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|1.1||
|formulasR1C1|object[][]|Represents the formula in R1C1-style notation.|1.2||
|hidden|bool|Represents if all cells of the current range are hidden. Read-only.|1.2||
|numberFormat|object[][]|Represents Excel's number format code for the given cell.|1.1||
|rowCount|int|Returns the total number of rows in the range. Read-only.|1.1||
|rowHidden|bool|Represents if all rows of the current range are hidden.|1.2||
|rowIndex|int|Returns the row number of the first cell in the range. Zero-indexed. Read-only.|1.1||
|text|object[][]|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|1.1||
|values|object[][]|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|1.1||

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|format|[RangeFormat](rangeformat.md)|Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties. Read-only.|1.1||
|sort|[RangeSort](rangesort.md)|The worksheet containing the current range. Read-only.|1.2||
|valueTypes|[RangeValueType](rangevaluetype.md)|Represents the type of data of each cell. Read-only.|1.1||
|visibleView|[RangeView](rangeview.md)|The visible part of the range. Read-only.|1.3||
|worksheet|[Worksheet](worksheet.md)|The worksheet containing the current range. Read-only.|1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clear(applyTo: string)](#clearapplyto-string)|void|Clear range values, format, fill, border, etc.|1.1|
|[delete(shift: string)](#deleteshift-string)|void|Deletes the cells associated with the range.|1.1|
|[getBoundingRect(anotherRange: Range or string)](#getboundingrectanotherrange-range-or-string)|[Range](range.md)|Gets the smallest range object that encompasses the given ranges. For example, the GetBoundingRect of "B2:C5" and "D10:E15" is "B2:E16".|1.1|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.|1.1|
|[getColumn(column: number)](#getcolumncolumn-number)|[Range](range.md)|Gets a column contained in the range.|1.1|
|[getEntireColumn()](#getentirecolumn)|[Range](range.md)|Gets an object that represents the entire column of the range.|1.1|
|[getEntireRow()](#getentirerow)|[Range](range.md)|Gets an object that represents the entire row of the range.|1.1|
|[getImage(height: number, width: number, fittingMode: ImageFittingMode)](#getimageheight-number-width-number-fittingmode-imagefittingmode)|[System.IO.Stream](system.io.stream.md)|Renders the range as a base64-encoded image by scaling the range to fit the specified dimensions.|1.3|
|[getIntersection(anotherRange: Range or string)](#getintersectionanotherrange-range-or-string)|[Range](range.md)|Gets the range object that represents the rectangular intersection of the given ranges.|1.1|
|[getIntersectionOrNull(anotherRange: Range or string)](#getintersectionornullanotherrange-range-or-string)|[Range](range.md)|Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.|1.3|
|[getLastCell()](#getlastcell)|[Range](range.md)|Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".|1.1|
|[getLastColumn()](#getlastcolumn)|[Range](range.md)|Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".|1.1|
|[getLastRow()](#getlastrow)|[Range](range.md)|Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".|1.1|
|[getOffsetRange(rowOffset: number, columnOffset: number)](#getoffsetrangerowoffset-number-columnoffset-number)|[Range](range.md)|Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an exception will be thrown.|1.1|
|[getRow(row: number)](#getrowrow-number)|[Range](range.md)|Gets a row contained in the range.|1.1|
|[getUsedRange(valuesOnly: [ApiSet(Version)](#getusedrangevaluesonly-apisetversion)|[Range](range.md)|Returns the used range of the given range object.|1.1|
|[insert(shift: string)](#insertshift-string)|[Range](range.md)|Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|
|[merge(across: bool)](#mergeacross-bool)|void|Merge the range cells into one region in the worksheet.|1.2|
|[select()](#select)|void|Selects the specified range in the Excel UI.|1.1|
|[unmerge()](#unmerge)|void|Unmerge the range cells into separate cells.|1.2|

## Method Details


### clear(applyTo: string)
Clear range values, format, fill, border, etc.

#### Syntax
```js
rangeObject.clear(applyTo);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|applyTo|string|Optional. Determines the type of clear action.|

#### Returns
void

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to clear the contents of the proxy range object.
    range.clear();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Cleared the selection (range object)');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### delete(shift: string)
Deletes the cells associated with the range.

#### Syntax
```js
rangeObject.delete(shift);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|shift|string|Specifies which way to shift the cells.|

#### Returns
void

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to delete the range object.
    range.delete();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Deleted the selection (range object)');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


### getBoundingRect(anotherRange: Range or string)
Gets the smallest range object that encompasses the given ranges. For example, the GetBoundingRect of "B2:C5" and "D10:E15" is "B2:E16".

#### Syntax
```js
rangeObject.getBoundingRect(anotherRange);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|anotherRange|Range or string|The range object or address or range name.|

#### Returns
[Range](range.md)

### getCell(row: number, column: number)
Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.

#### Syntax
```js
rangeObject.getCell(row, column);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|row|number|Row number of the cell to be retrieved. Zero-indexed.|
|column|number|Column number of the cell to be retrieved. Zero-indexed.|

#### Returns
[Range](range.md)

### getColumn(column: number)
Gets a column contained in the range.

#### Syntax
```js
rangeObject.getColumn(column);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|column|number|Column number of the range to be retrieved. Zero-indexed.|

#### Returns
[Range](range.md)

### getEntireColumn()
Gets an object that represents the entire column of the range.

#### Syntax
```js
rangeObject.getEntireColumn();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getEntireRow()
Gets an object that represents the entire row of the range.

#### Syntax
```js
rangeObject.getEntireRow();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getImage(height: number, width: number, fittingMode: ImageFittingMode)
Renders the range as a base64-encoded image by scaling the range to fit the specified dimensions.

#### Syntax
```js
rangeObject.getImage(height, width, fittingMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|height|number|Optional. (Optional) The desired height of the resulting image.|
|width|number|Optional. (Optional) The desired width of the resulting image.|
|fittingMode|ImageFittingMode|Optional. (Optional) The method used to scale the chart to the specified to the specified dimensions (if both height and width are set)."|

#### Returns
[System.IO.Stream](system.io.stream.md)

### getIntersection(anotherRange: Range or string)
Gets the range object that represents the rectangular intersection of the given ranges.

#### Syntax
```js
rangeObject.getIntersection(anotherRange);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|anotherRange|Range or string|The range object or range address that will be used to determine the intersection of ranges.|

#### Returns
[Range](range.md)

### getIntersectionOrNull(anotherRange: Range or string)
Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.

#### Syntax
```js
rangeObject.getIntersectionOrNull(anotherRange);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|anotherRange|Range or string|The range object or range address that will be used to determine the intersection of ranges.|

#### Returns
[Range](range.md)

### getLastCell()
Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".

#### Syntax
```js
rangeObject.getLastCell();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getLastColumn()
Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".

#### Syntax
```js
rangeObject.getLastColumn();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getLastRow()
Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".

#### Syntax
```js
rangeObject.getLastRow();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getOffsetRange(rowOffset: number, columnOffset: number)
Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an exception will be thrown.

#### Syntax
```js
rangeObject.getOffsetRange(rowOffset, columnOffset);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|rowOffset|number|The number of rows (positive, negative, or 0) by which the range is to be offset. Positive values are offset downward, and negative values are offset upward.|
|columnOffset|number|The number of columns (positive, negative, or 0) by which the range is to be offset. Positive values are offset to the right, and negative values are offset to the left.|

#### Returns
[Range](range.md)

### getRow(row: number)
Gets a row contained in the range.

#### Syntax
```js
rangeObject.getRow(row);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|row|number|Row number of the range to be retrieved. Zero-indexed.|

#### Returns
[Range](range.md)

### getUsedRange(valuesOnly: [ApiSet(Version)
Returns the used range of the given range object.

#### Syntax
```js
rangeObject.getUsedRange(valuesOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|valuesOnly|[ApiSet(Version|Considers only cells with values as used cells.|

#### Returns
[Range](range.md)

### insert(shift: string)
Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.

#### Syntax
```js
rangeObject.insert(shift);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|shift|string|Specifies which way to shift the cells.|

#### Returns
[Range](range.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert a page break after the selected text.
    range.insertBreak('page', 'After');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Inserted a page break after the selected text.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert a content control around the selected text,
    // and create a proxy content control object. We'll update the properties
    // on the content control.
    var myContentControl = range.insertContentControl();
    myContentControl.tag = "Customer-Address";
    myContentControl.title = "Enter Customer Address Here:";
    myContentControl.style = "Normal";
    myContentControl.insertText("One Microsoft Way, Redmond, WA 98052", 'replace');
    myContentControl.cannotEdit = true;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Wrapped a content control around the selected text.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert base64 encoded .docx at the beginning of the range.
    // You'll need to implement getBase64() to make this work.
    range.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Added base64 encoded text to the beginning of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert HTML in to the beginning of the range.
    range.insertHtml('<strong>This is text inserted with range.insertHtml()</strong>', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('HTML added to the beginning of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert OOXML in to the beginning of the range.
    range.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('OOXML added to the beginning of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

*Additional information* 
Read [Create better add-ins for Word with Office Open XML](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx) for guidance on working with OOXML.


#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert the paragraph after the range.
    range.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Paragraph added to the end of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert the paragraph at the end of the range.
    range.insertText('New text inserted into the range.', Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Text added to the end of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
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

### merge(across: bool)
Merge the range cells into one region in the worksheet.

#### Syntax
```js
rangeObject.merge(across);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|across|bool|Optional. Set true to merge cells in each row of the specified range as separate merged cells. The default value is false.|

#### Returns
void

### select()
Selects the specified range in the Excel UI.

#### Syntax
```js
rangeObject.select();
```

#### Parameters
None

#### Returns
void

### unmerge()
Unmerge the range cells into separate cells.

#### Syntax
```js
rangeObject.unmerge();
```

#### Parameters
None

#### Returns
void
