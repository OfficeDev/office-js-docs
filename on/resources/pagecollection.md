# PageCollection Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

Represents a collection of Pages.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[Page[]](page.md)|A collection of page objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getById(id: string)](#getbyidid-string)|[Page](page.md)|Gets a page by its identifier.|
|[getByTitle(title: string)](#getbytitletitle-string)|[PageCollection](pagecollection.md)|Gets the collection of Pages with specified title.|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[Page](page.md)|Gets an Page object by its index in the collection. Read-only.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getById(id: string)
Gets a page by its identifier.

#### Syntax
```js
pageCollectionObject.getById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|Required. A content control identifier.|

#### Returns
[Page](page.md)

### getByTitle(title: string)
Gets the collection of Pages with specified title.

#### Syntax
```js
pageCollectionObject.getByTitle(title);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|title|string|title of the Pages to search.|

#### Returns
[PageCollection](pagecollection.md)

### getItem(index: number or string)
Gets an Page object by its index in the collection. Read-only.

#### Syntax
```js
pageCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number or string|A number or an id that identifies the index location of a Page object.|

#### Returns
[Page](page.md)

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
