# Setting Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Represents a setting of the add-in.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|key|string|Gets the key of the setting. Read only. Read-only.|[1.4](../reqset/word-requirement.md)|
|value|object|Gets or sets the value of the setting.|[1.4](../reqset/word-requirement.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the setting.|[1.4](../reqset/word-requirement.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../reqset/word-requirement.md)|

## Method Details


### delete()
Deletes the setting.

#### Syntax
```js
settingObject.delete();
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
