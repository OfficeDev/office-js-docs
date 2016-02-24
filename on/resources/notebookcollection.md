# NotebookCollection Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

Represents a collection of Notebooks.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[Notebook[]](notebook.md)|A collection of notebook objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getById(id: string)](#getbyidid-string)|[Notebook](notebook.md)|Gets a notebook by its identifier.|
|[getByName(name: string)](#getbynamename-string)|[NotebookCollection](notebookcollection.md)|Gets the collection of Notebooks with specified name.|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[Notebook](notebook.md)|Gets an notebook object by its index in the collection. Read-only.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getById(id: string)
Gets a notebook by its identifier.

#### Syntax
```js
notebookCollectionObject.getById(id);
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
notebookCollectionObject.getByName(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the Notebook to search.|

#### Returns
[NotebookCollection](notebookcollection.md)

### getItem(index: number or string)
Gets an notebook object by its index in the collection. Read-only.

#### Syntax
```js
notebookCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number or string|A number or an id that identifies the index location of a notebook object.|

#### Returns
[Notebook](notebook.md)

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
