# Chart Object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

Represents a chart object in a workbook.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|height|double|Represents the height, in points, of the chart object.|
|id|string|Gets a chart based on its position in the collection. Read-only.|
|left|double|The distance, in points, from the left side of the chart to the worksheet origin.|
|name|string|Represents the name of a chart object.|
|top|double|Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
|width|double|Represents the width, in points, of the chart object.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|axes|[ChartAxes](chartaxes.md)|Represents chart axes. Read-only.|
|dataLabels|[ChartDataLabels](chartdatalabels.md)|Represents the datalabels on the chart. Read-only.|
|format|[ChartAreaFormat](chartareaformat.md)|Encapsulates the format properties for the chart area. Read-only.|
|legend|[ChartLegend](chartlegend.md)|Represents the legend for the chart. Read-only.|
|series|[ChartSeriesCollection](chartseriescollection.md)|Represents either a single series or collection of series in the chart. Read-only.|
|title|[ChartTitle](charttitle.md)|Represents the title of the specified chart, including the text, visibility, position and formating of the title. Read-only.|
|worksheet|[Worksheet](worksheet.md)|The worksheet containing the current chart. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Deletes the chart object.|
|[getImage(height: number, width: number, fittingMode: ImageFittingMode)](#getimageheight-number-width-number-fittingmode-imagefittingmode):new:|[System.IO.Stream](system.io.stream.md)|Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[setData(sourceData: Range, seriesBy: string)](#setdatasourcedata-range-seriesby-string)|void|Resets the source data for the chart.|
|[setPosition(startCell: Range or string, endCell: Range or string)](#setpositionstartcell-range-or-string-endcell-range-or-string)|void|Positions the chart relative to cells on the worksheet.|

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

_**Note**: This is a proposed feature that is still under design phase and hence not yet available as part of the product. The specification is being made available for community review and feedback. The final design may change. Help us make this feature better by providing your feedback [here](https://github.com/OfficeDev/office-js-docs/issues/new?title=ExcelJs-1.2-OpenSpec-chart-getImage)._

#### Syntax
```js
chartObject.getImage(height, width, fittingMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
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
|:---------------|:--------|:----------|
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
|:---------------|:--------|:----------|
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
|:---------------|:--------|:----------|
|startCell|Range or string|The start cell. This is where the chart will be moved to. The start cell is the top-left or top-right cell, depending on the user's right-to-left display settings.|
|endCell|Range or string|Optional. (Optional) The end cell. If specified, the chart's width and height will be set to fully cover up this cell/range.|

#### Returns
void
