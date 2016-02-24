# Paragraphs Object (JavaScript API for OneNote)

_Applies to: OneNote Online_

Represents a collection of Paragraphs.

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getById(id: string)](#getbyidid-string)|[Paragraph](paragraph.md)|Gets a paragraph by its identifier.|
|[getItem(index: number)](#getitemindex-number)|[Paragraph](paragraph.md)|Gets an Paragraph object by its index in the collection.|

## Method Details


### getById(id: string)
Gets a paragraph by its identifier.

#### Syntax
```js
paragraphsObject.getById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|Required. A content control identifier.|

#### Returns
[Paragraph](paragraph.md)

### getItem(index: number)
Gets an Paragraph object by its index in the collection.

#### Syntax
```js
paragraphsObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|A number that identifies the index location of a Paragraph object.|

#### Returns
[Paragraph](paragraph.md)
