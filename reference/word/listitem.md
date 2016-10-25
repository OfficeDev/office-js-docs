# ListItem Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Represents the paragraph list item format.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|level|int|Gets or sets the level of the item in the list.|[1.3](../reqset/word-requirement.md)|
|listString|string|Gets the list item bullet, number or picture as a string. Read-only.|[1.3](../reqset/word-requirement.md)|
|siblingIndex|int|Gets the list item order number in relation to its siblings. Read-only.|[1.3](../reqset/word-requirement.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getAncestor(parentOnly: bool)](#getancestorparentonly-bool)|[Paragraph](paragraph.md)|Gets the list item parent, or the closest ancestor if the parent does not exist.|[1.3](../reqset/word-requirement.md)|
|[getDescendants(directChildrenOnly: bool)](#getdescendantsdirectchildrenonly-bool)|[ParagraphCollection](paragraphcollection.md)|Gets all descendant list items of the list item.|[1.3](../reqset/word-requirement.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../reqset/word-requirement.md)|

## Method Details


### getAncestor(parentOnly: bool)
Gets the list item parent, or the closest ancestor if the parent does not exist.

#### Syntax
```js
listItemObject.getAncestor(parentOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|parentOnly|bool|Optional. Optional. Specified only the list item's parent will be returned. The default is false that specifies to get the lowest ancestor.|

#### Returns
[Paragraph](paragraph.md)

### getDescendants(directChildrenOnly: bool)
Gets all descendant list items of the list item.

#### Syntax
```js
listItemObject.getDescendants(directChildrenOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|directChildrenOnly|bool|Optional. Optional. Specified only the list item's direct children will be returned. The default is false that indicates to get all descendant items.|

#### Returns
[ParagraphCollection](paragraphcollection.md)

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
