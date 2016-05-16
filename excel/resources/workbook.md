# Workbook Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Workbook is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc.

## Properties

None

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|application|[Application](application.md)|Represents Excel application instance that contains this workbook. Read-only.|1.1||
|bindings|[BindingCollection](bindingcollection.md)|Represents a collection of bindings that are part of the workbook. Read-only.|1.1||
|functions|[Functions](functions.md)|Represents Excel application instance that contains this workbook. Read-only.|1.2||
|names|[NamedItemCollection](nameditemcollection.md)|Represents a collection of workbook scoped named items (named ranges and constants). Read-only.|1.1||
|pivotTables|[PivotTableCollection](pivottablecollection.md)|Returns sheet protection object for a worksheet. Read-only.|1.3||
|tables|[TableCollection](tablecollection.md)|Represents a collection of tables associated with the workbook. Read-only.|1.1||
|worksheets|[WorksheetCollection](worksheetcollection.md)|Represents a collection of worksheets associated with the workbook. Read-only.|1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|Gets the currently selected range from the workbook.|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### getSelectedRange()
Gets the currently selected range from the workbook.

#### Syntax
```js
workbookObject.getSelectedRange();
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
