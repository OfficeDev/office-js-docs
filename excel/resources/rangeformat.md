# RangeFormat Object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

A format object encapsulating the range's font, fill, borders, alignment, and other properties.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|columnWidth|double|Gets or sets the width of all colums within the range. If the column widths are not uniform, null will be returned.|
|horizontalAlignment|string|Represents the horizontal alignment for the specified object.|
|rowHeight|double|Gets or sets the height of all rows in the range. If the row heights are not uniform null will be returned.|
|verticalAlignment|string|Represents the vertical alignment for the specified object.|
|wrapText|bool|Indicates if Excel wraps the text in the object. A null value indicates that the entire range doesn't have uniform wrap setting|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|borders|[RangeBorderCollection](rangebordercollection.md)|Collection of border objects that apply to the overall range selected Read-only.|
|fill|[RangeFill](rangefill.md)|Returns the fill object defined on the overall range. Read-only.|
|font|[RangeFont](rangefont.md)|Returns the font object defined on the overall range selected Read-only.|
|protection|[FormatProtection](formatprotection.md)|Returns the format protection object for a range. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[autofitColumns()](#autofitcolumns)|void|Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.|
|[autofitRows()](#autofitrows)|void|Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

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
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
