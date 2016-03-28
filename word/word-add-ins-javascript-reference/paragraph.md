# Paragraph Object (JavaScript API for Word)

_Applies to: Word 2016, Word for iPad, Word for Mac 2016, Office 2016_

Represents a single paragraph in a selection, range, content control, or document body.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|listLevel|int|Gets or sets the list level of the paragraph.|
|outlineLevel|int|Gets or sets the outline level for the paragraph.|
|style|string|Gets or sets the style used for the paragraph. This is the name of the pre-installed or custom style.|
|text|string|Gets the text of the paragraph. Read-only.|



## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|alignment|[Alignment](alignment.md)|Gets or sets the alignment for a paragraph. The value can  be 'left', 'centered', 'right', or 'justified'.|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Gets the collection of content control objects that are in the paragraph. Read-only.|
|firstLineIndent|[float](float.md)|Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.|
|font|[Font](font.md)|Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties. Read-only.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Gets the collection of inlinePicture objects that are in the paragraph. The collection does not include floating images. Read-only.|
|leftIndent|[float](float.md)|Gets or sets the left indent value, in points, for the paragraph.|
|lineSpacing|[float](float.md)|Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.|
|lineUnitAfter|[float](float.md)|Gets or sets the amount of spacing, in grid lines. after the paragraph.|
|lineUnitBefore|[float](float.md)|Gets or sets the amount of spacing, in grid lines, before the paragraph.|
|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the paragraph. Returns null if there isn't a parent content control. Read-only.|
|rightIndent|[float](float.md)|Gets or sets the right indent value, in points, for the paragraph.|
|spaceAfter|[float](float.md)|Gets or sets the spacing, in points, after the paragraph.|
|spaceBefore|[float](float.md)|Gets or sets the spacing, in points, before the paragraph.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Clears the contents of the paragraph object. The user can perform the undo operation on the cleared content.|
|[delete()](#delete)|void|Deletes the paragraph and its content from the document.|
|[getHtml()](#gethtml)|string|Gets the HTML representation of the paragraph object.|
|[getOoxml()](#getooxml)|string|Gets the Office Open XML (OOXML) representation of the paragraph object.|
|[getRange(rangeLocation: RangeLocation)](#getrangerangelocation-rangelocation)|[Range](range.md)|Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Wraps the paragraph object with a rich text content control.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Inserts a document into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts HTML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Inserts a picture into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts OOXML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Inserts text into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|Performs a search with the specified searchOptions on the scope of the paragraph object. The search results are a collection of range objects.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Selects and navigates the Word UI to the paragraph.|
|[splitTextRanges(delimiters: string[], trimDelimiters: bool, trimWhitespace: bool)](#splittextrangesdelimiters-string-trimdelimiters-bool-trimwhitespace-bool)|[RangeCollection](rangecollection.md)|Splits the paragraph into child ranges by using delimiters.|

## Method Details


### clear()
Clears the contents of the paragraph object. The user can perform the undo operation on the cleared content.

#### Syntax
```js
paragraphObject.clear();
```

#### Parameters
None

#### Returns
void

### delete()
Deletes the paragraph and its content from the document.

#### Syntax
```js
paragraphObject.delete();
```

#### Parameters
None

#### Returns
void

### getHtml()
Gets the HTML representation of the paragraph object.

#### Syntax
```js
paragraphObject.getHtml();
```

#### Parameters
None

#### Returns
string

### getOoxml()
Gets the Office Open XML (OOXML) representation of the paragraph object.

#### Syntax
```js
paragraphObject.getOoxml();
```

#### Parameters
None

#### Returns
string

### getRange(rangeLocation: RangeLocation)
Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.

#### Syntax
```js
paragraphObject.getRange(rangeLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rangeLocation|RangeLocation|Optional. Optional. The range location can be 'Whole', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
paragraphObject.insertBreak(breakType, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|breakType|BreakType|Required. The break type to add to the document.|
|insertLocation|InsertLocation|Required. The value can be 'Before' or 'After'.|

#### Returns
void

### insertContentControl()
Wraps the paragraph object with a rich text content control.

#### Syntax
```js
paragraphObject.insertContentControl();
```

#### Parameters
None

#### Returns
[ContentControl](contentcontrol.md)

### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Inserts a document into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
paragraphObject.insertFileFromBase64(base64File, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64File|string|Required. The base64 encoded content of a .docx file.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### insertHtml(html: string, insertLocation: InsertLocation)
Inserts HTML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
paragraphObject.insertHtml(html, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|Required. The HTML to be inserted in the paragraph.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Inserts a picture into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
paragraphObject.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Required. The base64 encoded image to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Inserts OOXML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
paragraphObject.insertOoxml(ooxml, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|ooxml|string|Required. The OOXML to be inserted in the paragraph.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
paragraphObject.insertParagraph(paragraphText, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|paragraphText|string|Required. The paragraph text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Before' or 'After'.|

#### Returns
[Paragraph](paragraph.md)

### insertText(text: string, insertLocation: InsertLocation)
Inserts text into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
paragraphObject.insertText(text, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|text|string|Required. Text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

### search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)
Performs a search with the specified searchOptions on the scope of the paragraph object. The search results are a collection of range objects.

#### Syntax
```js
paragraphObject.search(searchText, searchOptions);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|searchText|string|Required. The search text.|
|searchOptions|ParamTypeStrings.SearchOptions|Optional. Optional. Options for the search.|

#### Returns
[SearchResultCollection](searchresultcollection.md)

### select(selectionMode: SelectionMode)
Selects and navigates the Word UI to the paragraph.

#### Syntax
```js
paragraphObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Optional. Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.|

#### Returns
void

### splitTextRanges(delimiters: string[], trimDelimiters: bool, trimWhitespace: bool)
Splits the paragraph into child ranges by using delimiters.

#### Syntax
```js
paragraphObject.splitTextRanges(delimiters, trimDelimiters, trimWhitespace);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|delimiters|string[]|Required. The delimiters as an array of strings.|
|trimDelimiters|bool|Optional. Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.|
|trimWhitespace|bool|Optional. Optional. Indicates whether to trim whitespace characters (spaces, tabs and column breaks) from the start and end of the ranges returned in the range collection. Default is false which indicates that whitespace characters at the start and end of the ranges are included in the range collection.|

#### Returns
[RangeCollection](rangecollection.md)
