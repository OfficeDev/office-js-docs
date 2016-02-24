# Paragraph Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

Represents a placeholder object for contents in an Outline. It can hold any one ParagraphType type of content. Paragraph is auto laid out.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID. Read-only.|
|type|string|Gets the ParagraphType of this Paragraph. Read-only. Possible values are: RichText, Image, Table, InkDrawing, InsertedFile, MediaFile, InkParagraph, Other.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|parentOutline|[Outline](outline.md)|Gets the parent Outline. Read-only.|
|parentParagraph|[Paragraph](paragraph.md)|Gets the parent Paragraph. It can be null if it is one of the top Paragraphs under Outline. Read-only.|
|subParagraphs|[ParagraphCollection](paragraphcollection.md)|Gets the child Paragraphs if it contains them. Only List and Table can have child Paragraphs. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getImage()](#getimage)|[Image](image.md)|Gets Image. Returns null if ParagraphType is not Image.|
|[getRichText()](#getrichtext)|[RichText](richtext.md)|Gets RichText. Returns null if ParagraphType is not RichText.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getImage()
Gets Image. Returns null if ParagraphType is not Image.

#### Syntax
```js
paragraphObject.getImage();
```

#### Parameters
None

#### Returns
[Image](image.md)

### getRichText()
Gets RichText. Returns null if ParagraphType is not RichText.

#### Syntax
```js
paragraphObject.getRichText();
```

#### Parameters
None

#### Returns
[RichText](richtext.md)

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
