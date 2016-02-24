# Pages Object (JavaScript API for OneNote)

_Applies to: OneNote Online_

Represents a collection of Pages.

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getById(id: string)](#getbyidid-string)|[Page](page.md)|Gets a page by its identifier.|
|[getByTitle(name: string)](#getbytitlename-string)|[Pages](pages.md)|Gets the collection of Pages with specified title.|
|[getItem(index: number)](#getitemindex-number)|[Page](page.md)|Gets an Page object by its index in the collection.|

## Method Details


### getById(id: string)
Gets a page by its identifier.

#### Syntax
```js
pagesObject.getById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|Required. A content control identifier.|

#### Returns
[Page](page.md)

### getByTitle(name: string)
Gets the collection of Pages with specified title.

#### Syntax
```js
pagesObject.getByTitle(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the Notebook to search.|

#### Returns
[Pages](pages.md)

### getItem(index: number)
Gets an Page object by its index in the collection.

#### Syntax
```js
pagesObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|A number that identifies the index location of a Page object.|

#### Returns
[Page](page.md)
