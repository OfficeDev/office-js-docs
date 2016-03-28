# Range Object (JavaScript API for Word)

_Applies to: Word 2016, Word for iPad, Word for Mac 2016, Office 2016_

Represents a contiguous area in a document.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|isEmpty|bool|Checks whether the range length is zero. Read-only.|
|style|string|Gets or sets the style used for the range. This is the name of the pre-installed or custom style.|
|text|string|Gets the text of the range. Read-only.|



## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Gets the collection of content control objects that are in the range. Read-only.|
|font|[Font](font.md)|Gets the text format of the range. Use this to get and set font name, size, color, and other properties. Read-only.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Gets the collection of inline picture objects that are in the range. Read-only.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Gets the collection of paragraph objects that are in the range. Read-only.|
|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the range. Returns null if there isn't a parent content control. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Clears the contents of the range object. The user can perform the undo operation on the cleared content.|
|[compareLocationWith(range: Range)](#comparelocationwithrange-range)|[LocationRelation](locationrelation.md)|Compares this range's location with another range's location.|
|[delete()](#delete)|void|Deletes the range and its content from the document.|
|[expandTo(range: Range)](#expandtorange-range)|void|Expands the range in either direction to cover another range.|
|[getHtml()](#gethtml)|string|Gets the HTML representation of the range object.|
|[getNextTextRange(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimWhitespace: bool)](#getnexttextrangedelimiters-string-multiparagraphs-bool-trimdelimiters-bool-trimwhitespace-bool)|[Range](range.md)|Gets the next text range by using delimiters.|
|[getOoxml()](#getooxml)|string|Gets the OOXML representation of the range object.|
|[getRange(rangeLocation: RangeLocation)](#getrangerangelocation-rangelocation)|[Range](range.md)|Clones the range, or gets the starting or ending point of the range as a new range.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Inserts a break at the specified location in the main document. The insertLocation value can be 'Replace', 'Before' or 'After'.|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Wraps the range object with a rich text content control.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Inserts a document at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Inserts a picture at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Inserts text at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|
|[intersectWith(range: Range)](#intersectwithrange-range)|void|Shrinks the range to the intersection of the range with another range.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|Performs a search with the specified searchOptions on the scope of the range object. The search results are a collection of range objects.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Selects and navigates the Word UI to the range.|
|[splitTextRanges(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimWhitespace: bool)](#splittextrangesdelimiters-string-multiparagraphs-bool-trimdelimiters-bool-trimwhitespace-bool)|[RangeCollection](rangecollection.md)|Splits the range into child ranges by using delimiters.|

## Method Details


### clear()
Clears the contents of the range object. The user can perform the undo operation on the cleared content.

#### Syntax
```js
rangeObject.clear();
```

#### Parameters
None

#### Returns
void

### compareLocationWith(range: Range)
Compares this range's location with another range's location.

#### Syntax
```js
rangeObject.compareLocationWith(range);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|range|Range|Required. The range to compare with this range.|

#### Returns
[LocationRelation](locationrelation.md)

### delete()
Deletes the range and its content from the document.

#### Syntax
```js
rangeObject.delete();
```

#### Parameters
None

#### Returns
void

### expandTo(range: Range)
Expands the range in either direction to cover another range.

#### Syntax
```js
rangeObject.expandTo(range);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|range|Range|Required. Another range.|

#### Returns
void

### getHtml()
Gets the HTML representation of the range object.

#### Syntax
```js
rangeObject.getHtml();
```

#### Parameters
None

#### Returns
string

### getNextTextRange(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimWhitespace: bool)
Gets the next text range by using delimiters.

#### Syntax
```js
rangeObject.getNextTextRange(delimiters, multiParagraphs, trimDelimiters, trimWhitespace);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|delimiters|string[]|Required. The delimiters as an array of strings.|
|multiParagraphs|bool|Optional. Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.|
|trimDelimiters|bool|Optional. Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.|
|trimWhitespace|bool|Optional. Optional. Indicates whether to trim whitespace characters (spaces, tabs and column breaks) from the start and end of the child ranges returned in the range collection. Default is false which indicates that whitespace characters at the start and end of the child ranges are included in the range collection.|

#### Returns
[Range](range.md)

### getOoxml()
Gets the OOXML representation of the range object.

#### Syntax
```js
rangeObject.getOoxml();
```

#### Parameters
None

#### Returns
string

### getRange(rangeLocation: RangeLocation)
Clones the range, or gets the starting or ending point of the range as a new range.

#### Syntax
```js
rangeObject.getRange(rangeLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rangeLocation|RangeLocation|Optional. Optional. The range location can be 'Whole', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Inserts a break at the specified location in the main document. The insertLocation value can be 'Replace', 'Before' or 'After'.

#### Syntax
```js
rangeObject.insertBreak(breakType, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|breakType|BreakType|Required. The break type to add.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Before' or 'After'.|

#### Returns
void

### insertContentControl()
Wraps the range object with a rich text content control.

#### Syntax
```js
rangeObject.insertContentControl();
```

#### Parameters
None

#### Returns
[ContentControl](contentcontrol.md)

### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Inserts a document at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
rangeObject.insertFileFromBase64(base64File, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64File|string|Required. The base64 encoded content of a .docx file.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|

#### Returns
[Range](range.md)

### insertHtml(html: string, insertLocation: InsertLocation)
Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
rangeObject.insertHtml(html, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|Required. The HTML to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|

#### Returns
[Range](range.md)

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Inserts a picture at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
rangeObject.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Required. The base64 encoded image to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|

#### Returns
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
rangeObject.insertOoxml(ooxml, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|ooxml|string|Required. The OOXML to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|

#### Returns
[Range](range.md)

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
rangeObject.insertParagraph(paragraphText, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|paragraphText|string|Required. The paragraph text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Before' or 'After'.|

#### Returns
[Paragraph](paragraph.md)

### insertText(text: string, insertLocation: InsertLocation)
Inserts text at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
rangeObject.insertText(text, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|text|string|Required. Text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|

#### Returns
[Range](range.md)

### intersectWith(range: Range)
Shrinks the range to the intersection of the range with another range.

#### Syntax
```js
rangeObject.intersectWith(range);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|range|Range|Required. Another range.|

#### Returns
void

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
Performs a search with the specified searchOptions on the scope of the range object. The search results are a collection of range objects.

#### Syntax
```js
rangeObject.search(searchText, searchOptions);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|searchText|string|Required. The search text.|
|searchOptions|ParamTypeStrings.SearchOptions|Optional. Optional. Options for the search.|

#### Returns
[SearchResultCollection](searchresultcollection.md)

### select(selectionMode: SelectionMode)
Selects and navigates the Word UI to the range.

#### Syntax
```js
rangeObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Optional. Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.|

#### Returns
void

### splitTextRanges(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimWhitespace: bool)
Splits the range into child ranges by using delimiters.

#### Syntax
```js
rangeObject.splitTextRanges(delimiters, multiParagraphs, trimDelimiters, trimWhitespace);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|delimiters|string[]|Required. The delimiters as an array of strings.|
|multiParagraphs|bool|Optional. Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.|
|trimDelimiters|bool|Optional. Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.|
|trimWhitespace|bool|Optional. Optional. Indicates whether to trim whitespace characters (spaces, tabs and column breaks) from the start and end of the child ranges returned in the range collection. Default is false which indicates that whitespace characters at the start and end of the child ranges are included in the range collection.|

#### Returns
[RangeCollection](rangecollection.md)
