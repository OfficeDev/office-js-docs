# RichText Object (JavaScript API for OneNote)

_Applies to: OneNote Online_

Represents RichText object.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|text|string|Gets text. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|parent|[Paragraph](paragraph.md)|Gets parent Paragraph. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getHtml()](#gethtml)|[String](string.md)|Gets html representing this RichText.|
|[insertHtml(html: string, insertLocation: string)](#inserthtmlhtml-string-insertlocation-string)|void|Inserts html to specified InsertLocation.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getHtml()
Gets html representing this RichText.

#### Syntax
```js
richTextObject.getHtml();
```

#### Parameters
None

#### Returns
[String](string.md)

### insertHtml(html: string, insertLocation: string)
Inserts html to specified InsertLocation.

#### Syntax
```js
richTextObject.insertHtml(html, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|Html to insert.|
|insertLocation|string|InsertLocation with respect to this RichText.  Possible values are: Before, After|

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
