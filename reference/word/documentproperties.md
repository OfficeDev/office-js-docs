# DocumentProperties Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Represents document properties.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|applicationName|string|Gets the application name of the document. Read only. Read-only.|[1.3](../reqset/word-requirement.md)|
|author|string|Gets or sets the author of the document.|[1.3](../reqset/word-requirement.md)|
|category|string|Gets or sets the category of the document.|[1.3](../reqset/word-requirement.md)|
|comments|string|Gets or sets the comments of the document.|[1.3](../reqset/word-requirement.md)|
|company|string|Gets or sets the company of the document.|[1.3](../reqset/word-requirement.md)|
|format|string|Gets or sets the format of the document.|[1.3](../reqset/word-requirement.md)|
|keywords|string|Gets or sets the keywords of the document.|[1.3](../reqset/word-requirement.md)|
|lastAuthor|string|Gets or sets the last author of the document.|[1.3](../reqset/word-requirement.md)|
|manager|string|Gets or sets the manager of the document.|[1.3](../reqset/word-requirement.md)|
|revisionNumber|string|Gets the revision number of the document. Read only. Read-only.|[1.3](../reqset/word-requirement.md)|
|security|int|Gets the security of the document. Read only. Read-only.|[1.3](../reqset/word-requirement.md)|
|subject|string|Gets or sets the subject of the document.|[1.3](../reqset/word-requirement.md)|
|template|string|Gets the template of the document. Read only. Read-only.|[1.3](../reqset/word-requirement.md)|
|title|string|Gets or sets the title of the document.|[1.3](../reqset/word-requirement.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|creationDate|[DateTime](datetime.md)|Gets the creation date of the document. Read only. Read-only.|[1.3](../reqset/word-requirement.md)|
|customProperties|[CustomPropertyCollection](custompropertycollection.md)|Gets the collection of custom properties of the document. Read only. Read-only.|[1.3](../reqset/word-requirement.md)|
|lastPrintDate|[DateTime](datetime.md)|Gets the last print date of the document. Read only. Read-only.|[1.3](../reqset/word-requirement.md)|
|lastSaveTime|[DateTime](datetime.md)|Gets the last save time of the document. Read only. Read-only.|[1.3](../reqset/word-requirement.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../reqset/word-requirement.md)|

## Method Details


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
