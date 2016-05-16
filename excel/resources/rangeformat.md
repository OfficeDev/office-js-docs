# RangeFormat Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

A format object encapsulating the range's font, fill, borders, alignment, and other properties.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|columnWidth|double|Gets or sets the width of all colums within the range. If the column widths are not uniform, null will be returned.|1.2||
|rowHeight|double|Gets or sets the height of all rows in the range. If the row heights are not uniform null will be returned.|1.2||
|wrapText|bool|Indicates if Excel wraps the text in the object. A null value indicates that the entire range doesn't have uniform wrap setting|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|borders|[RangeBorderCollection](rangebordercollection.md)|Collection of border objects that apply to the overall range selected Read-only.|1.1||
|fill|[RangeFill](rangefill.md)|Returns the fill object defined on the overall range. Read-only.|1.1||
|font|[RangeFont](rangefont.md)|Returns the font object defined on the overall range selected Read-only.|1.1||
|horizontalAlignment|string|Represents the horizontal alignment for the specified object.|1.1||
|protection|[FormatProtection](formatprotection.md)|Returns the format protection object for a range. Read-only.|1.2||
|verticalAlignment|string|Represents the vertical alignment for the specified object.|1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[autofitColumns()](#autofitcolumns)|void|Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.|1.2|
|[autofitRows()](#autofitrows)|void|Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.|1.2|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### autofitColumns()
Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.

#### Syntax
```js
rangeFormatObject.autofitColumns();
```

#### Parameters
None

#### Returns
void

### autofitRows()
Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.

#### Syntax
```js
rangeFormatObject.autofitRows();
```

#### Parameters
None

#### Returns
void

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
