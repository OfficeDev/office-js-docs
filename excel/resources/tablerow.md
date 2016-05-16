# TableRow Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents a row in a table.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|index|int|Returns the index number of the row within the rows collection of the table. Zero-indexed. Read-only.|1.1||
|values|object[][]|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the row from the table.|1.1|
|[getRange()](#getrange)|[Range](range.md)|Returns the range object associated with the entire row.|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### delete()
Deletes the row from the table.

#### Syntax
```js
tableRowObject.delete();
```

#### Parameters
None

#### Returns
void

### getRange()
Returns the range object associated with the entire row.

#### Syntax
```js
tableRowObject.getRange();
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
