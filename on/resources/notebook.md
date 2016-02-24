# Notebook Object (JavaScript API for OneNote)

_Applies to: OneNote Online_

Represents the Notebook in OneNote.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID. Read-only.|
|name|string|Gets the name. Read-only.|



## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getChildSectionGroups(recursive: bool)](#getchildsectiongroupsrecursive-bool)|[SectionGroups](sectiongroups.md)|Gets the child SectionGroups of this Notebook.|
|[getChildSections(recursive: bool)](#getchildsectionsrecursive-bool)|[Sections](sections.md)|Gets the child Sections of this Notebook.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getChildSectionGroups(recursive: bool)
Gets the child SectionGroups of this Notebook.

#### Syntax
```js
notebookObject.getChildSectionGroups(recursive);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|recursive|bool|Boolean value to indicate whether to retrieve all child SectionGroups or immediate child SectionGroups. |

#### Returns
[SectionGroups](sectiongroups.md)

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
