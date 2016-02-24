# Image Object (JavaScript API for OneNote)

_Applies to: OneNote Online_

Represents the Image object.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|description|string|Gets Image description. Read-only.|
|hyperlink|string|Gets the hyperlink.|



## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|height|[float](float.md)|Gets height of the image layout. Read-only.|
|parent|[Paragraph](paragraph.md)|Gets the parent Paragraph of this Image. Read-only.|
|width|[float](float.md)|Gets width of the image layout.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[changeDescription(description: string)](#changedescriptiondescription-string)|void|Gets width of the image layout.|
|[getBase64ImageSrc()](#getbase64imagesrc)|string|Gets width of the image layout.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### changeDescription(description: string)
Gets width of the image layout.

#### Syntax
```js
imageObject.changeDescription(description);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|description|string|Description for this Image.|

#### Returns
void

### getBase64ImageSrc()
Gets width of the image layout.

#### Syntax
```js
imageObject.getBase64ImageSrc();
```

#### Parameters
None

#### Returns
string

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
