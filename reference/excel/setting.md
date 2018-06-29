# Setting Object (JavaScript API for Excel)

Setting represents a key-value pair of a setting persisted to the document (per file per add-in). These custom key-value pair can be used to store state or lifecycle information needed by the content or task-pane add-in. Note that settings are persisted in the document and hence it is not a place to store any sensitive or protected information such as user information and password.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|key|string|Returns the key that represents the id of the Setting. Read-only.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|value|object|Represents the value stored for this setting.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the setting.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

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
