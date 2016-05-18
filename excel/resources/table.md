# Table Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents an Excel table.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|highlightFirstColumn|bool|Indicates whether the first column contains special formatting.|1.3||
|highlightLastColumn|bool|Indicates whether the last column contains special formatting.|1.3||
|id|int|Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed. Read-only.|1.1||
|name|string|Name of the table.|1.1||
|showBandedColumns|bool|Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.|1.3||
|showBandedRows|bool|Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.|1.3||
|showFilterButton|bool|Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.|1.3||
|showHeaders|bool|Indicates whether the header row is visible or not. This value can be set to show or remove the header row.|1.1||
|showTotals|bool|Indicates whether the total row is visible or not. This value can be set to show or remove the total row.|1.1||
|style|string|Constant value that represents the Table style. Possible values are: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|columns|[TableColumnCollection](tablecolumncollection.md)|Represents a collection of all the columns in the table. Read-only.|1.1||
|rows|[TableRowCollection](tablerowcollection.md)|Represents a collection of all the rows in the table. Read-only.|1.1||
|sort|[TableSort](tablesort.md)|Represents the sorting for the table. Read-only.|1.2||
|worksheet|[Worksheet](worksheet.md)|The worksheet containing the current table. Read-only.|1.2||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clearFilters()](#clearfilters)|void|Clears all the filters currently applied on the table.|1.2|
|[convertToRange()](#converttorange)|[Range](range.md)|Converts the table into a normal range of cells. All data is preserved.|1.2|
|[delete()](#delete)|void|Deletes the table.|1.1|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Gets the range object associated with the data body of the table.|1.1|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Gets the range object associated with header row of the table.|1.1|
|[getImage(height: number, width: number, fittingMode: ImageFittingMode)](#getimageheight-number-width-number-fittingmode-imagefittingmode)|[System.IO.Stream](system.io.stream.md)|Renders the table as a base64-encoded image by scaling the table to fit the specified dimensions.|1.3|
|[getRange()](#getrange)|[Range](range.md)|Gets the range object associated with the entire table.|1.1|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Gets the range object associated with totals row of the table.|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|
|[reapplyFilters()](#reapplyfilters)|void|Reapplies all the filters currently on the table.|1.2|

## Method Details


### clearFilters()
Clears all the filters currently applied on the table.

#### Syntax
```js
tableObject.clearFilters();
```

#### Parameters
None

#### Returns
void

### convertToRange()
Converts the table into a normal range of cells. All data is preserved.

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

### getImage(height: number, width: number, fittingMode: ImageFittingMode)
Renders the table as a base64-encoded image by scaling the table to fit the specified dimensions.

#### Syntax
```js
tableObject.getImage(height, width, fittingMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|height|number|Optional. (Optional) The desired height of the resulting image.|
|width|number|Optional. (Optional) The desired width of the resulting image.|
|fittingMode|ImageFittingMode|Optional. (Optional) The method used to scale the chart to the specified to the specified dimensions (if both height and width are set)."|

#### Returns
[System.IO.Stream](system.io.stream.md)

#### Example
```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.worksheets.getItem("Table1").charts.getItem("Table1"); 
    var image = chart.getImage();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

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
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

### reapplyFilters()
Reapplies all the filters currently on the table.

#### Syntax
```js
tableObject.reapplyFilters();
```

#### Parameters
None

#### Returns
void
