# PageContentCollection Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

Represents a collection of PageContents.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[PageContent[]](pagecontent.md)|A collection of pageContent objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getById(id: string)](#getbyidid-string)|[PageContent](pagecontent.md)|Gets a page content by its identifier.|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[PageContent](pagecontent.md)|Gets an PageContent object by its index in the collection. Read-only.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getById(id: string)
Gets a page content by its identifier.

#### Syntax
```js
pageContentCollectionObject.getById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|Required. A content control identifier.|

#### Returns
[PageContent](pagecontent.md)

### getItem(index: number or string)
Gets an PageContent object by its index in the collection. Read-only.

#### Syntax
```js
pageContentCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number or string|A number or an id that identifies the index location of a PageContent object.|

#### Returns
[PageContent](pagecontent.md)

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
