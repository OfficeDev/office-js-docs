# Sections Object (JavaScript API for OneNote)

_Applies to: OneNote Online_

Represents a collection of Sections.

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getById(id: string)](#getbyidid-string)|[Section](section.md)|Gets a section by its identifier.|
|[getByName(name: string)](#getbynamename-string)|[Sections](sections.md)|Gets the collection of Sections with specified name.|
|[getItem(index: number)](#getitemindex-number)|[Section](section.md)|Gets an section object by its index in the collection.|

## Method Details


### getById(id: string)
Gets a section by its identifier.

#### Syntax
```js
sectionsObject.getById(id);
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
sectionsObject.getByName(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the Notebook to search.|

#### Returns
[Sections](sections.md)

### getItem(index: number)
Gets an section object by its index in the collection.

#### Syntax
```js
sectionsObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|A number that identifies the index location of a section object.|

#### Returns
[Section](section.md)
