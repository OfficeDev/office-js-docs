# Table Object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

Represents an Excel table.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|int|Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed. Read-only.|
|name|string|Name of the table.|
|showHeaders|bool|Indicates whether the header row is visible or not. This value can be set to show or remove the header row.|
|showTotals|bool|Indicates whether the total row is visible or not. This value can be set to show or remove the total row.|
|style|string|Constant value that represents the Table style. Possible values are: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|columns|[TableColumnCollection](tablecolumncollection.md)|Represents a collection of all the columns in the table. Read-only.|
|rows|[TableRowCollection](tablerowcollection.md)|Represents a collection of all the rows in the table. Read-only.|
|**sort**|[TableSort](tablesort.md)|Represents the sorting for the table. Read-only.|
|worksheet|[Worksheet](worksheet.md)|The worksheet containing the current table. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|**[convertToRange()](#converttorange)**|[Range](range.md)|Converts the table into a normal range of cells. All data is preserved.|
|[delete()](#delete)|void|Deletes the table.|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Gets the range object associated with the data body of the table.|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Gets the range object associated with header row of the table.|
|[getRange()](#getrange)|[Range](range.md)|Gets the range object associated with the entire table.|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Gets the range object associated with totals row of the table.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### convertToRange()
Converts the table into a normal range of cells. All data is preserved.

_**Note**: This is a proposed feature that is still under design phase and hence not yet available as part of the product. The specification is being made available for community review and feedback. The final design may change. Help us make this feature better by providing your feedback [here](https://github.com/OfficeDev/office-js-docs/issues/new?title=ExcelJs-1.2-OpenSpec-table-convertToRange)._

#### Syntax
```js
tableObject.convertToRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### delete()
Deletes the table.

#### Syntax
```js
tableObject.delete();
```

#### Parameters
None

#### Returns
void

### getDataBodyRange()
Gets the range object associated with the data body of the table.

#### Syntax
```js
tableObject.getDataBodyRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getHeaderRowRange()
Gets the range object associated with header row of the table.

#### Syntax
```js
tableObject.getHeaderRowRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getRange()
Gets the range object associated with the entire table.

#### Syntax
```js
tableObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getTotalRowRange()
Gets the range object associated with totals row of the table.

#### Syntax
```js
tableObject.getTotalRowRange();
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
