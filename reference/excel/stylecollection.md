# StyleCollection Object (JavaScript API for Excel)

Represents a collection of all the styles.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[Style[]](style.md)|A collection of style objects. Read-only.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(name: string)](#addname-string)|void|Adds a new style to the collection.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[Style](style.md)|Gets a style by name.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(name: string)
Adds a new style to the collection.

#### Syntax
```js
styleCollectionObject.add(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the style to be added.|

#### Returns
void

### getItem(name: string)
Gets a style by name.

#### Syntax
```js
styleCollectionObject.getItem(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the style to be retrieved.|

#### Returns
[Style](style.md)
