# ContentControl Object (JavaScript API for Word)

_Applies to: Word 2016, Word for iPad, Word for Mac 2016, Office 2016_

Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|cannotDelete|bool|Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.|
|cannotEdit|bool|Gets or sets a value that indicates whether the user can edit the contents of the content control.|
|color|string|Gets or sets the color of the content control. Color is set in '#RRGGBB' format or by using the color name.|
|placeholderText|string|Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.|
|removeWhenEdited|bool|Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.|
|style|string|Gets or sets the style used for the content control. This is the name of the pre-installed or custom style.|
|tag|string|Gets or sets a tag to identify a content control.|
|text|string|Gets the text of the content control. Read-only.|
|title|string|Gets or sets the title for a content control.|



## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|appearance|[ContentControlAppearance](contentcontrolappearance.md)|Gets or sets the appearance of the content control. The value can be 'boundingBox', 'tags' or 'hidden'.|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Gets the collection of content control objects in the content control. Read-only.|
|font|[Font](font.md)|Gets the text format of the content control. Use this to get and set font name, size, color, and other properties. Read-only.|
|id|[uint](uint.md)|Gets an integer that represents the content control identifier. Read-only.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Gets the collection of inlinePicture objects in the content control. The collection does not include floating images. Read-only.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Get the collection of paragraph objects in the content control. Read-only.|
|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the content control. Returns null if there isn't a parent content control. Read-only.|
|type|[ContentControlType](contentcontroltype.md)|Gets the content control type. Only rich text content controls are supported currently. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Clears the contents of the content control. The user can perform the undo operation on the cleared content.|
|[delete(keepContent: bool)](#deletekeepcontent-bool)|void|Deletes the content control and its content. If keepContent is set to true, the content is not deleted.|
|[getHtml()](#gethtml)|string|Gets the HTML representation of the content control object.|
|[getOoxml()](#getooxml)|string|Gets the Office Open XML (OOXML) representation of the content control object.|
|[getRange(rangeLocation: RangeLocation)](#getrangerangelocation-rangelocation)|[Range](range.md)|Gets the whole content control, or the starting or ending point of the content control, as a range.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Inserts a break at the specified location in the main document. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Inserts a document into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts HTML into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Inserts an inline picture into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts OOXML into the content control at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Inserts text into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|Performs a search with the specified searchOptions on the scope of the content control object. The search results are a collection of range objects.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Selects the content control. This causes Word to scroll to the selection.|
|[splitTextRanges(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimWhitespace: bool)](#splittextrangesdelimiters-string-multiparagraphs-bool-trimdelimiters-bool-trimwhitespace-bool)|[RangeCollection](rangecollection.md)|Splits the content control into child ranges by using delimiters.|

## Method Details


### clear()
Clears the contents of the content control. The user can perform the undo operation on the cleared content.

#### Syntax
```js
contentControlObject.clear();
```

#### Parameters
None

#### Returns
void

### delete(keepContent: bool)
Deletes the content control and its content. If keepContent is set to true, the content is not deleted.

#### Syntax
```js
contentControlObject.delete(keepContent);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|keepContent|bool|Required. Indicates whether the content should be deleted with the content control. If keepContent is set to true, the content is not deleted.|

#### Returns
void

### getHtml()
Gets the HTML representation of the content control object.

#### Syntax
```js
contentControlObject.getHtml();
```

#### Parameters
None

#### Returns
string

### getOoxml()
Gets the Office Open XML (OOXML) representation of the content control object.

#### Syntax
```js
contentControlObject.getOoxml();
```

#### Parameters
None

#### Returns
string

### getRange(rangeLocation: RangeLocation)
Gets the whole content control, or the starting or ending point of the content control, as a range.

#### Syntax
```js
contentControlObject.getRange(rangeLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rangeLocation|RangeLocation|Optional. Optional. The range location can be 'Whole', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Inserts a break at the specified location in the main document. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
contentControlObject.insertBreak(breakType, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|breakType|BreakType|Required. Type of break.|
|insertLocation|InsertLocation|Required. The value can be 'Start', 'End', 'Before' or 'After'.|

#### Returns
void

### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Inserts a document into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
contentControlObject.insertFileFromBase64(base64File, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64File|string|Required. The base64 encoded content of a .docx file.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### insertHtml(html: string, insertLocation: InsertLocation)
Inserts HTML into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
contentControlObject.insertHtml(html, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|Required. The HTML to be inserted in to the content control.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Inserts an inline picture into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
contentControlObject.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Required. The base64 encoded image to be inserted in the content control.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Inserts OOXML into the content control at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
contentControlObject.insertOoxml(ooxml, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|ooxml|string|Required. The OOXML to be inserted in to the content control.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
contentControlObject.insertParagraph(paragraphText, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|paragraphText|string|Required. The paragrph text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Start', 'End', 'Before' or 'After'.|

#### Returns
[Paragraph](paragraph.md)

### insertText(text: string, insertLocation: InsertLocation)
Inserts text into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
contentControlObject.insertText(text, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|text|string|Required. The text to be inserted in to the content control.|
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
Performs a search with the specified searchOptions on the scope of the content control object. The search results are a collection of range objects.

#### Syntax
```js
contentControlObject.search(searchText, searchOptions);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|searchText|string|Required. The search text.|
|searchOptions|ParamTypeStrings.SearchOptions|Optional. Optional. Options for the search.|

#### Returns
[SearchResultCollection](searchresultcollection.md)

### select(selectionMode: SelectionMode)
Selects the content control. This causes Word to scroll to the selection.

#### Syntax
```js
contentControlObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Optional. Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.|

#### Returns
void

### splitTextRanges(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimWhitespace: bool)
Splits the content control into child ranges by using delimiters.

#### Syntax
```js
contentControlObject.splitTextRanges(delimiters, multiParagraphs, trimDelimiters, trimWhitespace);
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
