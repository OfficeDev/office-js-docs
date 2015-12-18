# TableColumn Object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

Represents a column in a table.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|int|Returns a unique key that identifies the column within the table. Read-only.|
|index|int|Returns the index number of the column within the columns collection of the table. Zero-indexed. Read-only.|
|name|string|Returns the name of the table column. Read-only.|
|values|object[][]|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Deletes the column from the table.|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Gets the range object associated with the data body of the column.|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Gets the range object associated with the header row of the column.|
|[getRange()](#getrange)|[Range](range.md)|Gets the range object associated with the entire column.|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Gets the range object associated with the totals row of the column.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

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
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
