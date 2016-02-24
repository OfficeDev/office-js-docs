# SectionGroupCollection Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

Represents the SectionGroup in OneNote.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[SectionGroup[]](sectiongroup.md)|A collection of sectionGroup objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getById(id: string)](#getbyidid-string)|[SectionGroup](sectiongroup.md)|Gets a section group by its identifier.|
|[getByName(name: string)](#getbynamename-string)|[SectionGroupCollection](sectiongroupcollection.md)|Gets the collection of SectionGroups with specified name.|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[SectionGroup](sectiongroup.md)|Gets an SectionGroup object by its index in the collection. Read-only.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getById(id: string)
Gets a section group by its identifier.

#### Syntax
```js
sectionGroupCollectionObject.getById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|Required. A content control identifier.|

#### Returns
[SectionGroup](sectiongroup.md)

### getByName(name: string)
Gets the collection of SectionGroups with specified name.

#### Syntax
```js
sectionGroupCollectionObject.getByName(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|name of the SectionGroup to search.|

#### Returns
[SectionGroupCollection](sectiongroupcollection.md)

### getItem(index: number or string)
Gets an SectionGroup object by its index in the collection. Read-only.

#### Syntax
```js
sectionGroupCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number or string|A number or an id that identifies the index location of a SectionGroup object.|

#### Returns
[SectionGroup](sectiongroup.md)

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
