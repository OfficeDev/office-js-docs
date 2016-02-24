# SectionGroup Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

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
|[getChildSectionGroups()](#getchildsectiongroups)|[SectionGroupCollection](sectiongroupcollection.md)|Gets the child SectionGroups.|
|[getChildSections(recursive: bool)](#getchildsectionsrecursive-bool)|[SectionCollection](sectioncollection.md)|Gets the collection of Sections.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getChildSectionGroups()
Gets the child SectionGroups.

#### Syntax
```js
sectionGroupObject.getChildSectionGroups();
```

#### Parameters
None

#### Returns
[SectionGroupCollection](sectiongroupcollection.md)

### getChildSections(recursive: bool)
Gets the collection of Sections.

#### Syntax
```js
sectionGroupObject.getChildSections(recursive);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|recursive|bool|Boolean value to indicate whether to retrieve all child sections or immediate child sections.|

#### Returns
[SectionCollection](sectioncollection.md)

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
