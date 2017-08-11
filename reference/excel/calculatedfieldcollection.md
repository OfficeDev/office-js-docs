# CalculatedFieldCollection Object (JavaScript API for Excel)

A collection of PivotField objects that represents all the calculated fields in the specified PivotTable report.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[CalculatedField[]](calculatedfield.md)|A collection of calculatedField objects. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(Name: string, Formula: string, UseStandardFormula: bool)](#addname-string-formula-string-usestandardformula-bool)|[PivotField](pivotfield.md)|Creates a new calculated field. Returns a PivotField object.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Returns the number of objects in the collection.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(nameOrIndex: object)](#getitemnameorindex-object)|[PivotField](pivotfield.md)|Returns a single object from a collection based on name or index.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[PivotField](pivotfield.md)|Returns a single object from a collection based on index.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(Name: string, Formula: string, UseStandardFormula: bool)
Creates a new calculated field. Returns a PivotField object.

#### Syntax
```js
calculatedFieldCollectionObject.add(Name, Formula, UseStandardFormula);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|Name|string|Name of the pivot field to be added.|
|Formula|string|Formula of the pivot field. |
|UseStandardFormula|bool|Optional. True evaluates field names using U.S. English settings; False evaluates names using the user's locale settings. Default is False.|

#### Returns
[PivotField](pivotfield.md)

### getCount()
Returns the number of objects in the collection.

#### Syntax
```js
calculatedFieldCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(nameOrIndex: object)
Returns a single object from a collection based on name or index.

#### Syntax
```js
calculatedFieldCollectionObject.getItem(nameOrIndex);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|nameOrIndex|object|Name or index of the item to be retrieved.|

#### Returns
[PivotField](pivotfield.md)

### getItemAt(index: number)
Returns a single object from a collection based on index.

#### Syntax
```js
calculatedFieldCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index of the item. Zero-indexed.|

#### Returns
[PivotField](pivotfield.md)
