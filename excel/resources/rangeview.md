# RangeView Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

RangeView represents an abstract of Range, including basic Range properties values, text, numberFormat, address.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|cellCount|int|Number of cells in the range. Read-only.|1.3||
|columnCount|int|Represents the total number of columns in the range. Read-only.|1.1||
|formulas|object[][]|Represents the formula in A1-style notation.|1.3||
|formulasLocal|object[][]|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|1.3||
|formulasR1C1|object[][]|Represents the formula in R1C1-style notation.|1.2||
|numberFormat|object[][]|Represents Excel's number format code for the given cell.|1.3||
|rowCount|int|Returns the total number of rows in the range. Read-only.|1.3||
|text|object[][]|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|1.3||
|values|object[][]|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|1.3||
|valueTypes|string[][]|Represents the type of data of each cell. Read-only. Possible values are: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.3||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|rows|[RangeViewCollection](rangeviewcollection.md)|Represents a collection of all the rows in the RangeView. Read-only.|1.3||


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|

## Method Details

