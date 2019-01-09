# CustomPropertyCollection Object (JavaScript API for Excel)

Contains the collection of customProperty objects.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[CustomProperty[]](customproperty.md)|A collection of customProperty objects. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(key: string, value: object)](#addkey-string-value-object)|[CustomProperty](customproperty.md)|Creates a new or sets an existing custom property.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[deleteAll()](#deleteall)|void|Deletes all custom properties in this collection.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Gets the count of custom properties.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[CustomProperty](customproperty.md)|Gets a custom property object by its key, which is case-insensitive. Throws if the custom property does not exist.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: string)](#getitemornullobjectkey-string)|[CustomProperty](customproperty.md)|Gets a custom property object by its key, which is case-insensitive. Returns a null object if the custom property does not exist.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(key: string, value: object)
Creates a new or sets an existing custom property.

#### Syntax
```js
customPropertyCollectionObject.add(key, value);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|string|Required. The custom property's key, which is case-insensitive.|
|value|object|Required. The custom property's value.|

#### Returns
[CustomProperty](customproperty.md)

### deleteAll()
Deletes all custom properties in this collection.

#### Syntax
```js
customPropertyCollectionObject.deleteAll();
```

#### Parameters
None

#### Returns
void

### getCount()
Gets the count of custom properties.

#### Syntax
```js
customPropertyCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(key: string)
Gets a custom property object by its key, which is case-insensitive. Throws if the custom property does not exist.

#### Syntax
```js
customPropertyCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|string|The key that identifies the custom property object.|

#### Returns
[CustomProperty](customproperty.md)

### getItemOrNullObject(key: string)
Gets a custom property object by its key, which is case-insensitive. Returns a null object if the custom property does not exist.

#### Syntax
```js
customPropertyCollectionObject.getItemOrNullObject(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|string|Required. The key that identifies the custom property object.|

#### Returns
[CustomProperty](customproperty.md)
