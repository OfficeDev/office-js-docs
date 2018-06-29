# PageLayout Object (JavaScript API for Excel)

Represents a collection of all the styles.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|blackAndWhite|bool|Gets or sets the worksheet's black and white print option.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|bottomMargin|double|Gets or sets the worksheet's bottom page margin to use for printing in points.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|centerHorizontally|bool|Gets or sets the worksheet's center horizontally flag. This flag determines whether the worksheet will be centered horizontally when it's printed.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|centerVertically|bool|Gets or sets the worksheet's center vertically flag. This flag determines whether the worksheet will be centered vertically when it's printed.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|draftMode|bool|Gets or sets the worksheet's draft mode option. If true the sheet will be printed without graphics.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|firstPageNumber|object|Gets or sets the worksheet's first page number to print. Null value represents "auto" page numbering.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|footerMargin|double|Gets or sets the worksheet's footer margin, in points, for use when printing.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|headerMargin|double|Gets or sets the worksheet's header margin, in points, for use when printing.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|leftMargin|double|Gets or sets the worksheet's left margin, in points, for use when printing.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|orientation|string|Gets or sets the worksheet's orientation of the page.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|paperSize|string|Gets or sets the worksheet's paper size of the page.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|printGridlines|bool|Gets or sets the worksheet's print gridlines flag. This flag determines whether gridlines will be printed or not.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|printHeadings|bool|Gets or sets the worksheet's print headings flag. This flag determines whether headings will be printed or not.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|rightMargin|double|Gets or sets the worksheet's right margin, in points, for use when printing.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|topMargin|double|Gets or sets the worksheet's top margin, in points, for use when printing.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|headersFooters|[HeaderFooterGroup](headerfootergroup.md)|Header and footer configuration for the worksheet. Read-only.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|printComments|[PrintComments](printcomments.md)|Gets or sets whether the worksheet's comments should be displayed when printing.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|printErrors|[PrintErrorType](printerrortype.md)|Gets or sets the worksheet's print errors option.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|printOrder|[PrintOrder](printorder.md)|Gets or sets the worksheet's page print order option. This specifies the order to use for processing the page number printed.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|zoom|[PageLayoutZoomOptions](pagelayoutzoomoptions.md)|Gets or sets the worksheet's print zoom options.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getPrintArea()](#getprintarea)|[Range](range.md)|Gets the range object representing the print area for the worksheet. If not set, this will return a null object.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|[getPrintAreaOrNullObject()](#getprintareaornullobject)|[Range](range.md)|Gets the range object representing the print area for the worksheet.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|[getPrintTitleColumns()](#getprinttitlecolumns)|[Range](range.md)|Gets the range object representing the title columns.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|[getPrintTitleColumnsOrNullObject()](#getprinttitlecolumnsornullobject)|[Range](range.md)|Gets the range object representing the title columns. If not set, this will return a null object.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|[getPrintTitleRows()](#getprinttitlerows)|[Range](range.md)|Gets the range object representing the title rows.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|[getPrintTitleRowsOrNullObject()](#getprinttitlerowsornullobject)|[Range](range.md)|Gets the range object representing the title rows. If not set, this will return a null object.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|[setPrintArea(printArea: Range or string)](#setprintareaprintarea-range-or-string)|void|Sets the worksheet's print area.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|[setPrintMargins(unit: PrintMarginUnit, marginOptions: PageLayoutMarginOptions)](#setprintmarginsunit-printmarginunit-marginoptions-pagelayoutmarginoptions)|void|Sets the worksheet's page margins with units.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|[setPrintTitleColumns(printTitleColumns: Range or string)](#setprinttitlecolumnsprinttitlecolumns-range-or-string)|void|Sets the columns that contain the cells to be repeated at the left of each page of the worksheet for printing.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|[setPrintTitleRows(printTitleRows: Range or string)](#setprinttitlerowsprinttitlerows-range-or-string)|void|Sets the rows that contain the cells to be repeated at the top of each page of the worksheet for printing.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getPrintArea()
Gets the range object representing the print area for the worksheet. If not set, this will return a null object.

#### Syntax
```js
pageLayoutObject.getPrintArea();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getPrintAreaOrNullObject()
Gets the range object representing the print area for the worksheet.

#### Syntax
```js
pageLayoutObject.getPrintAreaOrNullObject();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getPrintTitleColumns()
Gets the range object representing the title columns.

#### Syntax
```js
pageLayoutObject.getPrintTitleColumns();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getPrintTitleColumnsOrNullObject()
Gets the range object representing the title columns. If not set, this will return a null object.

#### Syntax
```js
pageLayoutObject.getPrintTitleColumnsOrNullObject();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getPrintTitleRows()
Gets the range object representing the title rows.

#### Syntax
```js
pageLayoutObject.getPrintTitleRows();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getPrintTitleRowsOrNullObject()
Gets the range object representing the title rows. If not set, this will return a null object.

#### Syntax
```js
pageLayoutObject.getPrintTitleRowsOrNullObject();
```

#### Parameters
None

#### Returns
[Range](range.md)

### setPrintArea(printArea: Range or string)
Sets the worksheet's print area.

#### Syntax
```js
pageLayoutObject.setPrintArea(printArea);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|printArea|Range or string|The range of the content to print.|

#### Returns
void

### setPrintMargins(unit: PrintMarginUnit, marginOptions: PageLayoutMarginOptions)
Sets the worksheet's page margins with units.

#### Syntax
```js
pageLayoutObject.setPrintMargins(unit, marginOptions);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|unit|PrintMarginUnit|Measurement unit for the margins provided.|
|marginOptions|PageLayoutMarginOptions|Margin values to set, margins not provided will remain unchanged.|

#### Returns
void

### setPrintTitleColumns(printTitleColumns: Range or string)
Sets the columns that contain the cells to be repeated at the left of each page of the worksheet for printing.

#### Syntax
```js
pageLayoutObject.setPrintTitleColumns(printTitleColumns);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|printTitleColumns|Range or string|The columns to be repeated to the left of each page, range must span the entire column to be valid.|

#### Returns
void

### setPrintTitleRows(printTitleRows: Range or string)
Sets the rows that contain the cells to be repeated at the top of each page of the worksheet for printing.

#### Syntax
```js
pageLayoutObject.setPrintTitleRows(printTitleRows);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|printTitleRows|Range or string|The rows to be repeated at the top of each page, range must span the entire row to be valid.|

#### Returns
void
