# SectionCollection Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

Represents a collection of Sections.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[Section[]](section.md)|A collection of section objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getById(id: string)](#getbyidid-string)|[Section](section.md)|Gets a section by its identifier.|
|[getByName(name: string)](#getbynamename-string)|[SectionCollection](sectioncollection.md)|Gets the collection of Sections with specified name.|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[Section](section.md)|Gets an section object by its index in the collection. Read-only.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getById(id: string)
Gets a section by its identifier.

#### Syntax
```js
sectionCollectionObject.getById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|Required. A content control identifier.|

#### Returns
[Section](section.md)

### getByName(name: string)
Gets the collection of Sections with specified name.

#### Syntax
```js
sectionCollectionObject.getByName(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the Notebook to search.|

#### Returns
[SectionCollection](sectioncollection.md)

### getItem(index: number or string)
Gets an section object by its index in the collection. Read-only.

#### Syntax
```js
sectionCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number or string|A number or an id that identifies the index location of a section object.|

#### Returns
[Section](section.md)

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
