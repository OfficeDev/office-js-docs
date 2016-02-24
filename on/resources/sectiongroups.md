# SectionGroups Object (JavaScript API for OneNote)

_Applies to: OneNote Online_

Represents the SectionGroup in OneNote.

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getById(id: string)](#getbyidid-string)|[SectionGroup](sectiongroup.md)|Gets a section group by its identifier.|
|[getByName(name: string)](#getbynamename-string)|[SectionGroups](sectiongroups.md)|Gets the SectionGroup with specified name.|
|[getItem(index: number)](#getitemindex-number)|[SectionGroup](sectiongroup.md)|Gets an SectionGroup object by its index in the collection.|

## Method Details


### getById(id: string)
Gets a section group by its identifier.

#### Syntax
```js
sectionGroupsObject.getById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|Required. A content control identifier.|

#### Returns
[SectionGroup](sectiongroup.md)

### getByName(name: string)
Gets the SectionGroup with specified name.

#### Syntax
```js
sectionGroupsObject.getByName(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|name of the SectionGroup to search.|

#### Returns
[SectionGroups](sectiongroups.md)

### getItem(index: number)
Gets an SectionGroup object by its index in the collection.

#### Syntax
```js
sectionGroupsObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|A number that identifies the index location of a SectionGroup object.|

#### Returns
[SectionGroup](sectiongroup.md)
