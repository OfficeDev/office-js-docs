# TableScopedCollection Object (JavaScript API for Excel)

Represents a scoped collection of tables. For each table its top-left corner is considered its anchor location and the tables are sorted top to bottom and then left to right.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[TableScoped[]](tablescoped.md)|A collection of tableScoped objects. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Gets the number of tables in the collection.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getFirst()](#getfirst)|[Table](table.md)|Gets the first table in the collection. The tables in the collection are sorted top to bottom and left to right, such that top left table is the first table in the collection.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[Table](table.md)|Gets a table by Name or ID.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCount()
Gets the number of tables in the collection.

#### Syntax
```js
tableScopedCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getFirst()
Gets the first table in the collection. The tables in the collection are sorted top to bottom and left to right, such that top left table is the first table in the collection.

#### Syntax
```js
tableScopedCollectionObject.getFirst();
```

#### Parameters
None

#### Returns
[Table](table.md)

### getItem(key: string)
Gets a table by Name or ID.

#### Syntax
```js
tableScopedCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|string|Name or ID of the table to be retrieved.|

#### Returns
[Table](table.md)
