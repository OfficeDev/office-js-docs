# Body Object (JavaScript API for Word)

_Applies to: Word 2016, Word for iPad, Word for Mac 2016, Office 2016_

Represents the body of a document or a section.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|style|string|Gets or sets the style used for the body. This is the name of the pre-installed or custom style.|
|text|string|Gets the text of the body. Use the insertText method to insert text. Read-only.|



## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Gets the collection of rich text content control objects that are in the body. Read-only.|
|font|[Font](font.md)|Gets the text format of the body. Use this to get and set font name, size, color, and other properties. Read-only.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Gets the collection of inlinePicture objects that are in the body. The collection does not include floating images. Read-only.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Gets the collection of paragraph objects that are in the body. Read-only.|
|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the body. Returns null if there isn't a parent content control. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Clears the contents of the body object. The user can perform the undo operation on the cleared content.|
|[getHtml()](#gethtml)|string|Gets the HTML representation of the body object.|
|[getOoxml()](#getooxml)|string|Gets the OOXML (Office Open XML) representation of the body object.|
|[getRange(rangeLocation: RangeLocation)](#getrangerangelocation-rangelocation)|[Range](range.md)|Gets the whole body, or the starting or ending point of the body, as a range.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Inserts a break at the specified location in the main document. The insertLocation value can be 'Start' or 'End'.|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Wraps the body object with a Rich Text content control.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Inserts a document into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Inserts a picture into the body at the specified location. The insertLocation value can be 'Start' or 'End'.|
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Start' or 'End'.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Inserts text into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|Performs a search with the specified searchOptions on the scope of the body object. The search results are a collection of range objects.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Selects the body and navigates the Word UI to it.|
|[splitTextRanges(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimWhitespace: bool)](#splittextrangesdelimiters-string-multiparagraphs-bool-trimdelimiters-bool-trimwhitespace-bool)|[RangeCollection](rangecollection.md)|Splits the body into child ranges by using delimiters.|

## Method Details


### clear()
Clears the contents of the body object. The user can perform the undo operation on the cleared content.

#### Syntax
```js
bodyObject.clear();
```

#### Parameters
None

#### Returns
void

### getHtml()
Gets the HTML representation of the body object.

#### Syntax
```js
bodyObject.getHtml();
```

#### Parameters
None

#### Returns
string

### getOoxml()
Gets the OOXML (Office Open XML) representation of the body object.

#### Syntax
```js
bodyObject.getOoxml();
```

#### Parameters
None

#### Returns
string

### getRange(rangeLocation: RangeLocation)
Gets the whole body, or the starting or ending point of the body, as a range.

#### Syntax
```js
bodyObject.getRange(rangeLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rangeLocation|RangeLocation|Optional. Optional. The range location can be 'Whole', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Inserts a break at the specified location in the main document. The insertLocation value can be 'Start' or 'End'.

#### Syntax
```js
bodyObject.insertBreak(breakType, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|breakType|BreakType|Required. The break type to add to the body.|
|insertLocation|InsertLocation|Required. The value can be 'Start' or 'End'.|

#### Returns
void

### insertContentControl()
Wraps the body object with a Rich Text content control.

#### Syntax
```js
bodyObject.insertContentControl();
```

#### Parameters
None

#### Returns
[ContentControl](contentcontrol.md)

### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Inserts a document into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
bodyObject.insertFileFromBase64(base64File, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64File|string|Required. The base64 encoded content of a .docx file.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### insertHtml(html: string, insertLocation: InsertLocation)
Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
bodyObject.insertHtml(html, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|Required. The HTML to be inserted in the document.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Inserts a picture into the body at the specified location. The insertLocation value can be 'Start' or 'End'.

#### Syntax
```js
bodyObject.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Required. The base64 encoded image to be inserted in the body.|
|insertLocation|InsertLocation|Required. The value can be 'Start' or 'End'.|

#### Returns
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
bodyObject.insertOoxml(ooxml, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|ooxml|string|Required. The OOXML to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Inserts a paragraph at the specified location. The insertLocation value can be 'Start' or 'End'.

#### Syntax
```js
bodyObject.insertParagraph(paragraphText, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|paragraphText|string|Required. The paragraph text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Start' or 'End'.|

#### Returns
[Paragraph](paragraph.md)

### insertText(text: string, insertLocation: InsertLocation)
Inserts text into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
bodyObject.insertText(text, insertLocation);
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
Performs a search with the specified searchOptions on the scope of the body object. The search results are a collection of range objects.

#### Syntax
```js
bodyObject.search(searchText, searchOptions);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|searchText|string|Required. The search text.|
|searchOptions|ParamTypeStrings.SearchOptions|Optional. Optional. Options for the search.|

#### Returns
[SearchResultCollection](searchresultcollection.md)

### select(selectionMode: SelectionMode)
Selects the body and navigates the Word UI to it.

#### Syntax
```js
bodyObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Optional. Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.|

#### Returns
void

### splitTextRanges(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimWhitespace: bool)
Splits the body into child ranges by using delimiters.

#### Syntax
```js
bodyObject.splitTextRanges(delimiters, multiParagraphs, trimDelimiters, trimWhitespace);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|delimiters|string[]|Required. The delimiters as an array of strings.|
|multiParagraphs|bool|Optional. Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.|
|trimDelimiters|bool|Optional. Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.|
|trimWhitespace|bool|Optional. Optional. Indicates whether to trim whitespace characters (spaces, tabs and column breaks) from the start and end of the ranges returned in the range collection. Default is false which indicates that whitespace characters at the start and end of the ranges are included in the range collection.|

#### Returns
[RangeCollection](rangecollection.md)
