# CustomProperty Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Represents a custom property.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|key|string|Gets the key of the custom property. Read only. Read-only.|[1.3](../reqset/word-requirement.md)|
|value|object|Gets or sets the value of the custom property.|[1.3](../reqset/word-requirement.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|type|[DocumentPropertyType](documentpropertytype.md)|Gets the value type of the custom property. Read only. Read-only.|[1.3](../reqset/word-requirement.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the custom property.|[1.3](../reqset/word-requirement.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../reqset/word-requirement.md)|

## Method Details


### delete()
Deletes the custom property.

#### Syntax
```js
customPropertyObject.delete();
```

#### Parameters
None

#### Returns
void

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
