# TableColumnCollection Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents a collection of all the columns that are part of the table.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|count|int|Returns the number of columns in the table. Read-only.|1.1||
|items|[TableColumn[]](tablecolumn.md)|A collection of tableColumn objects. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(index: number, values: (boolean or string or number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableColumn](tablecolumn.md)|Adds a new column to the table.|1.1|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[TableColumn](tablecolumn.md)|Gets a column object by Name or ID.|1.1|
|[getItemAt(index: number)](#getitematindex-number)|[TableColumn](tablecolumn.md)|Gets a column based on its position in the collection.|1.1|
|[getItemOrNull(key: number or string)](#getitemornullkey-number-or-string)|[TableColumn](tablecolumn.md)|Gets a column object by Name or ID. If the column does not exist, the returned object's isNull property will be true.|1.3|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### add(index: number, values: (boolean or string or number)[][])
Adds a new column to the table.

#### Syntax
```js
tableColumnCollectionObject.add(index, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|Specifies the relative position of the new column. The previous column at this position is shifted to the right. The index value should be equal to or less than the last column's index value, so it cannot be used to append a column at the end of the table. Zero-indexed.|
|values|(boolean or string or number)[][]|Optional. A 2-dimensional array of unformatted values of the table column.|

#### Returns
[TableColumn](tablecolumn.md)

### getItem(key: number or string)
Gets a column object by Name or ID.

#### Syntax
```js
tableColumnCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|number or string| Column Name or ID.|

#### Returns
[TableColumn](tablecolumn.md)

### getItemAt(index: number)
Gets a column based on its position in the collection.

#### Syntax
```js
tableColumnCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[TableColumn](tablecolumn.md)

### getItemOrNull(key: number or string)
Gets a column object by Name or ID. If the column does not exist, the returned object's isNull property will be true.

#### Syntax
```js
tableColumnCollectionObject.getItemOrNull(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|number or string| Column Name or ID.|

#### Returns
[TableColumn](tablecolumn.md)

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
