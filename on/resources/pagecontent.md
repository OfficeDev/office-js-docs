# PageContent Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

Represents a placeholder object for top leve content under page. These contents can be asigned x,y position in a page.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID. Read-only.|
|type|string|Gets the Type of the Content. Read-only. Possible values are: Outline, Image, Ink, InsertedFile, MediaFile, Other.|
|xPosition|int|Gets or Sets the x position of this content.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|get;|[{]({.md)|Gets or Sets the y position of this content.|
|parent|[Page](page.md)|Gets the parent Page. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Deletes this PageContent.|
|[getImage()](#getimage)|[Image](image.md)|Gets the Image. It will return null if the PageContentType is not Image.|
|[getOutline()](#getoutline)|[Outline](outline.md)|Gets the Outline. It will return null if the PageContentType is not Outline.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### delete()
Deletes this PageContent.

#### Syntax
```js
pageContentObject.delete();
```

#### Parameters
None

#### Returns
void

### getImage()
Gets the Image. It will return null if the PageContentType is not Image.

#### Syntax
```js
pageContentObject.getImage();
```

#### Parameters
None

#### Returns
[Image](image.md)

### getOutline()
Gets the Outline. It will return null if the PageContentType is not Outline.

#### Syntax
```js
pageContentObject.getOutline();
```

#### Parameters
None

#### Returns
[Outline](outline.md)

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
