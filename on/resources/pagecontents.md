# PageContents Object (JavaScript API for OneNote)

_Applies to: OneNote Online_

Represents a collection of PageContents.

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getById(id: string)](#getbyidid-string)|[PageContent](pagecontent.md)|Gets a page content by its identifier.|
|[getItem(index: number)](#getitemindex-number)|[PageContent](pagecontent.md)|Gets an PageContent object by its index in the collection.|

## Method Details


### getById(id: string)
Gets a page content by its identifier.

#### Syntax
```js
pageContentsObject.getById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|Required. A content control identifier.|

#### Returns
[PageContent](pagecontent.md)

### getItem(index: number)
Gets an PageContent object by its index in the collection.

#### Syntax
```js
pageContentsObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|A number that identifies the index location of a PageContent object.|

#### Returns
[PageContent](pagecontent.md)
