# PageContent Object (JavaScript API for OneNote)

_Applies to: OneNote Online_

Represents a placeholder object for top leve content under page. These contents can be asigned x,y position in a page.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID. Read-only.|
|type|string|Gets the Type of the Content. Read-only. Possible values are: Outline, Image, Ink, InsertedFile, MediaFile, Other.|



## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|parent|[Page](page.md)|Gets the parent Page. Read-only.|
|position|[Position](position.md)|Gets or Sets the x,y position of this content.|

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
