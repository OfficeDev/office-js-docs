# Document Object (JavaScript API for Word)

_Applies to: Word 2016, Word for iPad, Word for Mac 2016, Office 2016_

The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|saved|bool|Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.|



## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|body|[Body](body.md)|Gets the body of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only.|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Gets the collection of content control objects that are in the current document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only.|
|sections|[SectionCollection](sectioncollection.md)|Gets the collection of section objects that are in the document. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getSelection()](#getselection)|[Range](range.md)|Gets the current selection of the document. Multiple selections are not supported.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[open()](#open)|void|Open the document.|
|[save()](#save)|void|Saves the document. This will use the Word default file naming convention if the document has not been saved before.|

## Method Details


### getSelection()
Gets the current selection of the document. Multiple selections are not supported.

#### Syntax
```js
documentObject.getSelection();
```

#### Parameters
None

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

### open()
Open the document.

#### Syntax
```js
documentObject.open();
```

#### Parameters
None

#### Returns
void

### save()
Saves the document. This will use the Word default file naming convention if the document has not been saved before.

#### Syntax
```js
documentObject.save();
```

#### Parameters
None

#### Returns
void
