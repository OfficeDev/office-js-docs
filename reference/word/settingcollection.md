# SettingCollection Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Contains the collection of [setting](setting.md) objects.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[Setting[]](setting.md)|A collection of setting objects. Read-only.|[1.4](../requirement-sets/word-api-requirement.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[deleteAll()](#deleteall)|void|Deletes all settings in this add-in.|[1.4](../requirement-sets/word-api-requirement.md)|
|[getCount()](#getcount)|int|Gets the count of settings.|[1.4](../requirement-sets/word-api-requirement.md)|
|[getItem(key: string)](#getitemkey-string)|[Setting](setting.md)|Gets a setting object by its key, which is case-sensitive. Throws if the setting does not exist.|[1.4](../requirement-sets/word-api-requirement.md)|
|[getItemOrNullObject(key: string)](#getitemornullobjectkey-string)|[Setting](setting.md)|Gets a setting object by its key, which is case-sensitive. Returns a null object if the setting does not exist.|[1.4](../requirement-sets/word-api-requirement.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/word-api-requirement.md)|
|[set(key: string, value: object)](#setkey-string-value-object)|[Setting](setting.md)|Creates or sets a setting.|[1.4](../requirement-sets/word-api-requirement.md)|

## Method Details


### deleteAll()
Deletes all settings in this add-in.

#### Syntax
```js
settingCollectionObject.deleteAll();
```

#### Parameters
None

#### Returns
void

### getCount()
Gets the count of settings.

#### Syntax
```js
settingCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(key: string)
Gets a setting object by its key, which is case-sensitive. Throws if the setting does not exist.

#### Syntax
```js
settingCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|string|The key that identifies the setting object.|

#### Returns
[Setting](setting.md)

### getItemOrNullObject(key: string)
Gets a setting object by its key, which is case-sensitive. Returns a null object if the setting does not exist.

#### Syntax
```js
settingCollectionObject.getItemOrNullObject(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|string|Required. The key that identifies the setting object.|

#### Returns
[Setting](setting.md)

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

### set(key: string, value: object)
Creates or sets a setting.

#### Syntax
```js
settingCollectionObject.set(key, value);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|string|Required. The setting's key, which is case-sensitive.|
|value|object|Required. The setting's value.|

#### Returns
[Setting](setting.md)
