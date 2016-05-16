# Chart Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents a chart object in a workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|height|double|Represents the height, in points, of the chart object.|1.1||
|id|string|Gets a chart based on its position in the collection. Read-only.|1.2||
|left|double|The distance, in points, from the left side of the chart to the worksheet origin.|1.1||
|name|string|Represents the name of a chart object.|1.1||
|top|double|Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|1.1||
|width|double|Represents the width, in points, of the chart object.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|axes|[ChartAxes](chartaxes.md)|Represents chart axes. Read-only.|1.1||
|dataLabels|[ChartDataLabels](chartdatalabels.md)|Represents the datalabels on the chart. Read-only.|1.1||
|format|[ChartAreaFormat](chartareaformat.md)|Encapsulates the format properties for the chart area. Read-only.|1.1||
|legend|[ChartLegend](chartlegend.md)|Represents the legend for the chart. Read-only.|1.1||
|series|[ChartSeriesCollection](chartseriescollection.md)|Represents either a single series or collection of series in the chart. Read-only.|1.1||
|title|[ChartTitle](charttitle.md)|Represents the title of the specified chart, including the text, visibility, position and formating of the title. Read-only.|1.1||
|worksheet|[Worksheet](worksheet.md)|The worksheet containing the current chart. Read-only.|1.2||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the chart object.|1.1|
|[getImage(height: number, width: number, fittingMode: ImageFittingMode)](#getimageheight-number-width-number-fittingmode-imagefittingmode)|[System.IO.Stream](system.io.stream.md)|Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.|1.2|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|
|[setData(sourceData: Range, seriesBy: string)](#setdatasourcedata-range-seriesby-string)|void|Resets the source data for the chart.|1.1|
|[setPosition(startCell: Range or string, endCell: Range or string)](#setpositionstartcell-range-or-string-endcell-range-or-string)|void|Positions the chart relative to cells on the worksheet.|1.1|

## Method Details


### delete()
Deletes the chart object.

#### Syntax
```js
chartObject.delete();
```

#### Parameters
None

#### Returns
void

### getImage(height: number, width: number, fittingMode: ImageFittingMode)
Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.

#### Syntax
```js
chartObject.getImage(height, width, fittingMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|height|number|Optional. (Optional) The desired height of the resulting image.|
|width|number|Optional. (Optional) The desired width of the resulting image.|
|fittingMode|ImageFittingMode|Optional. (Optional) The method used to scale the chart to the specified to the specified dimensions (if both height and width are set)."|

#### Returns
[System.IO.Stream](system.io.stream.md)

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

### setData(sourceData: Range, seriesBy: string)
Resets the source data for the chart.

#### Syntax
```js
chartObject.setData(sourceData, seriesBy);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|sourceData|Range|The Range object corresponding to the source data.|
|seriesBy|string|Optional. Specifies the way columns or rows are used as data series on the chart. Can be one of the following: Auto (default), Rows, Columns.|

#### Returns
void

### setPosition(startCell: Range or string, endCell: Range or string)
Positions the chart relative to cells on the worksheet.

#### Syntax
```js
chartObject.setPosition(startCell, endCell);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|startCell|Range or string|The start cell. This is where the chart will be moved to. The start cell is the top-left or top-right cell, depending on the user's right-to-left display settings.|
|endCell|Range or string|Optional. (Optional) The end cell. If specified, the chart's width and height will be set to fully cover up this cell/range.|

#### Returns
void
