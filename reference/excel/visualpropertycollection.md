# VisualPropertyCollection Object (JavaScript API for Excel)

Represents a collection of visual object properties.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[VisualProperty[]](visualproperty.md)|A collection of visualProperty objects. Read-only.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Returns the number of properties in the collection.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(index: number)](#getitemindex-number)|[VisualProperty](visualproperty.md)|Returns a property at given index|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[VisualProperty](visualproperty.md)|Returns a property at given index|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCount()
Returns the number of properties in the collection.

#### Syntax
```js
visualPropertyCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(index: number)
Returns a property at given index

#### Syntax
```js
visualPropertyCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Represents the index in property array.|

#### Returns
[VisualProperty](visualproperty.md)

### getItemAt(index: number)
Returns a property at given index

#### Syntax
```js
visualPropertyCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Represents the index in property array.|

#### Returns
[VisualProperty](visualproperty.md)
