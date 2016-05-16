# TableColumn Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents a column in a table.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|int|Returns a unique key that identifies the column within the table. Read-only.|1.1||
|index|int|Returns the index number of the column within the columns collection of the table. Zero-indexed. Read-only.|1.1||
|name|string|Returns the name of the table column. Read-only.|1.1||
|values|object[][]|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|filter|[Filter](filter.md)|Retrieve the filter applied to the column. Read-only.|1.2||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the column from the table.|1.1|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Gets the range object associated with the data body of the column.|1.1|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Gets the range object associated with the header row of the column.|1.1|
|[getRange()](#getrange)|[Range](range.md)|Gets the range object associated with the entire column.|1.1|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Gets the range object associated with the totals row of the column.|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### delete()
Deletes the column from the table.

#### Syntax
```js
tableColumnObject.delete();
```

#### Parameters
None

#### Returns
void

### getDataBodyRange()
Gets the range object associated with the data body of the column.

#### Syntax
```js
tableColumnObject.getDataBodyRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getHeaderRowRange()
Gets the range object associated with the header row of the column.

#### Syntax
```js
tableColumnObject.getHeaderRowRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getRange()
Gets the range object associated with the entire column.

#### Syntax
```js
tableColumnObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getTotalRowRange()
Gets the range object associated with the totals row of the column.

#### Syntax
```js
tableColumnObject.getTotalRowRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

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
