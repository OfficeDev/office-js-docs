# Notebooks Object (JavaScript API for OneNote)

_Applies to: OneNote Online_

Represents a collection of Notebooks.

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getById(id: string)](#getbyidid-string)|[Notebook](notebook.md)|Gets a notebook by its identifier.|
|[getByName(name: string)](#getbynamename-string)|[Notebooks](notebooks.md)|Gets the collection of Notebooks with specified name.|
|[getItem(index: number)](#getitemindex-number)|[Notebook](notebook.md)|Gets an notebook object by its index in the collection.|

## Method Details


### getById(id: string)
Gets a notebook by its identifier.

#### Syntax
```js
notebooksObject.getById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|Required. A content control identifier.|

#### Returns
[Notebook](notebook.md)

### getByName(name: string)
Gets the collection of Notebooks with specified name.

#### Syntax
```js
notebooksObject.getByName(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the Notebook to search.|

#### Returns
[Notebooks](notebooks.md)

### getItem(index: number)
Gets an notebook object by its index in the collection.

#### Syntax
```js
notebooksObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|A number that identifies the index location of a notebook object.|

#### Returns
[Notebook](notebook.md)
