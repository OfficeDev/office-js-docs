# ContentControlCollection Object (JavaScript API for Word)

_Applies to: Word 2016, Word for iPad, Word for Mac 2016, Office 2016_

Contains a collection of ContentControl objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[ContentControl[]](contentcontrol.md)|A collection of contentControl objects. Read-only.|



## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getById(id: number)](#getbyidid-number)|[ContentControl](contentcontrol.md)|Gets a content control by its identifier.|
|[getByTag(tag: string)](#getbytagtag-string)|[ContentControlCollection](contentcontrolcollection.md)|Gets the content controls that have the specified tag.|
|[getByTitle(title: string)](#getbytitletitle-string)|[ContentControlCollection](contentcontrolcollection.md)|Gets the content controls that have the specified title.|
|[getItem(index: number)](#getitemindex-number)|[ContentControl](contentcontrol.md)|Gets a content control by its index in the collection.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getById(id: number)
Gets a content control by its identifier.

#### Syntax
```js
contentControlCollectionObject.getById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|number|Required. A content control identifier.|

#### Returns
[ContentControl](contentcontrol.md)

### getByTag(tag: string)
Gets the content controls that have the specified tag.

#### Syntax
```js
contentControlCollectionObject.getByTag(tag);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|tag|string|Required. A tag set on a content control.|

#### Returns
[ContentControlCollection](contentcontrolcollection.md)

### getByTitle(title: string)
Gets the content controls that have the specified title.

#### Syntax
```js
contentControlCollectionObject.getByTitle(title);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|title|string|Required. The title of a content control.|

#### Returns
[ContentControlCollection](contentcontrolcollection.md)

### getItem(index: number)
Gets a content control by its index in the collection.

#### Syntax
```js
contentControlCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|The index|

#### Returns
[ContentControl](contentcontrol.md)

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
