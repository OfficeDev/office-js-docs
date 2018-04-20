# AutoFilter Object (JavaScript API for Excel)
AutoFilter Object
Represents an AutoFilter on a worksheet or table

> **Note:** When using AutoFilter with dates, the format should be consistent with English date separators ("/") instead of local settings ("."). A valid date would be "2/2/2007", whereas "2.2.2007" is invalid.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
| filters |FilterCollection[]|A collection of filter objects.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
| range | Range[]| Range object that represents the range to which the specified AutoFilter applies. |[beta](../requirement-sets/excel-api-requirement-sets.md)|
| filterMode | bool| Returns True if the worksheet is in the AutoFilter filter mode. Read-only |[beta](../requirement-sets/excel-api-requirement-sets.md)|
| sort | RangeSort[]| Gets the sort column or columns, and sort order for the AutoFilter on the range. |[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|applyFilter()|void|Applies the Autofilter object.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|showgitAllData()|void|Displays all the data returned by the AutoFilter object..|[beta](../requirement-sets/excel-api-requirement-sets.md)|
