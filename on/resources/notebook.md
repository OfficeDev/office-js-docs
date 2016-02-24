# Notebook Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

Represents the Notebook in OneNote.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID. Read-only.|
|name|string|Gets the name. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getChildSectionGroups()](#getchildsectiongroups)|[SectionGroupCollection](sectiongroupcollection.md)|Gets the child SectionGroups of this Notebook.|
|[getChildSections(recursive: bool)](#getchildsectionsrecursive-bool)|[SectionCollection](sectioncollection.md)|Gets the child Sections of this Notebook.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getChildSectionGroups()
Gets the child SectionGroups of this Notebook.

#### Syntax
```js
notebookObject.getChildSectionGroups();
```

#### Parameters
None

#### Returns
[SectionGroupCollection](sectiongroupcollection.md)

### getChildSections(recursive: bool)
Gets the child Sections of this Notebook.

#### Syntax
```js
notebookObject.getChildSections(recursive);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|recursive|bool|Boolean value to indicate whether to retrieve all child sections or immediate child sections. |

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
