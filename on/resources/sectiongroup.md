# SectionGroup Object (JavaScript API for OneNote)

_Applies to: OneNote Online_

Represents the SectionGroup in OneNote.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID. Read-only.|
|name|string|Gets the name. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|parentNotebook|[Notebook](notebook.md)|Gets the Notebook of this SectionGroup. It can never be null. Read-only.|
|parentSectionGroup|[SectionGroup](sectiongroup.md)|Gets the parent SectionGroup if it exists. Otherwise returns null which indicates this is under Notebook. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getChildSectionGroups(recursive: bool)](#getchildsectiongroupsrecursive-bool)|[SectionGroups](sectiongroups.md)|todo|
|[getChildSections(name: bool)](#getchildsectionsname-bool)|[Sections](sections.md)|Gets the collection of Notebooks with specified name.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getChildSectionGroups(recursive: bool)
todo

#### Syntax
```js
sectionGroupObject.getChildSectionGroups(recursive);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|recursive|bool|todo.|

#### Returns
[SectionGroups](sectiongroups.md)

### getChildSections(name: bool)
Gets the collection of Notebooks with specified name.

#### Syntax
```js
sectionGroupObject.getChildSections(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|bool|Name of the Notebook to search.|

#### Returns
[Sections](sections.md)

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
