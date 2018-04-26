# Chart Object (JavaScript API for Excel)

Represents a chart object in a workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|categoryLabelLevel|int|Returns or sets a ChartCategoryLabelLevel enumeration constant referring to the level of where the category labels are being sourced from. ReadWrite.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|chartType|string|Represents the type of the chart. Possible values are: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc..|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|colorScheme|int|Returns or sets an integer that represents the color scheme for the chart. ReadWrite.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|displayBlanksAs|string|Returns or sets the way that blank cells are plotted on a chart. ReadWrite. Possible values are: NotPlotted, Zero, Interplotted.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|height|double|Represents the height, in points, of the chart object.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|The unique id of chart. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|left|double|The distance, in points, from the left side of the chart to the worksheet origin.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Represents the name of a chart object.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|plotBy|string|Returns or sets the way columns or rows are used as data series on the chart. ReadWrite. Possible values are: Rows, Columns.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|plotVisibleOnly|bool|True if only visible cells are plotted. False if both visible and hidden cells are plotted. ReadWrite.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|roundedCorners|bool|True if the chart area of the chart has rounded corners. ReadWrite.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|seriesNameLevel|int|Returns or sets a ChartSeriesNameLevel enumeration constant referring to|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|showAllFieldButtons|bool|Represents whether to display all field buttons on a PivotChart.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|showAxisFieldButtons|bool|Represents whether to display axis field buttons on a PivotChart. The ShowAxisFieldButtons property corresponds to the Show Axis Field Buttons command on the Field Buttons drop-down list of the Analyze tab, which is available when a PivotChart is selected.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|showDataLabelsOverMaximum|bool|Represents whether to to show the data labels when the value is greater than the maximum value on the value axis. If value axis became smaller than the size of data points, you can use this property to set whether to show the data labels. This property applies to 2-D charts only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|showLegendFieldButtons|bool|Represents whether to display legend field buttons on a PivotChart.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|showReportFilterFieldButtons|bool|Represents whether to display report filter field buttons on a PivotChart.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|showValueFieldButtons|bool|Represents whether to display show value field buttons on a PivotChart.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|style|int|Returns or sets the chart style for the chart. ReadWrite.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|top|double|Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|width|double|Represents the width, in points, of the chart object.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|axes|[ChartAxes](chartaxes.md)|Represents chart axes. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|dataLabels|[ChartDataLabels](chartdatalabels.md)|Represents the datalabels on the chart. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|format|[ChartAreaFormat](chartareaformat.md)|Encapsulates the format properties for the chart area. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|legend|[ChartLegend](chartlegend.md)|Represents the legend for the chart. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|plotArea|[ChartPlotArea](chartplotarea.md)|Represents the plotArea for the chart. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|series|[ChartSeriesCollection](chartseriescollection.md)|Represents either a single series or collection of series in the chart. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|title|[ChartTitle](charttitle.md)|Represents the title of the specified chart, including the text, visibility, position and formating of the title. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|worksheet|[Worksheet](worksheet.md)|The worksheet containing the current chart. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the chart object.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getImage(height: number, width: number, fittingMode: string)](#getimageheight-number-width-number-fittingmode-string)|[System.IO.Stream](system.io.stream.md)|Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[setData(sourceData: object, seriesBy: string)](#setdatasourcedata-object-seriesby-string)|void|Resets the source data for the chart.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[setPosition(startCell: object, endCell: object)](#setpositionstartcell-object-endcell-object)|void|Positions the chart relative to cells on the worksheet.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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

#### Examples
```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.delete();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getImage(height: number, width: number, fittingMode: string)
Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.

#### Syntax
```js
chartObject.getImage(height, width, fittingMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|height|number|Optional. (Optional) The desired height of the resulting image.|
|width|number|Optional. (Optional) The desired width of the resulting image.|
|fittingMode|string|Optional. (Optional) The method used to scale the chart to the specified to the specified dimensions (if both height and width are set)."  Possible values are: Fit, FitAndCenter, Fill|

#### Returns
[System.IO.Stream](system.io.stream.md)

#### Examples
```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var image = chart.getImage();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```





### setData(sourceData: object, seriesBy: string)
Resets the source data for the chart.

#### Syntax
```js
chartObject.setData(sourceData, seriesBy);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|sourceData|object|The Range object corresponding to the source data.|
|seriesBy|string|Optional. Specifies the way columns or rows are used as data series on the chart. Can be one of the following: Auto (default), Rows, Columns.  Possible values are: Auto, Columns, Rows|

#### Returns
void

#### Examples

Set the `sourceData` to be "A1:B4" and `seriesBy` to be "Columns"

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var sourceData = "A1:B4";
	chart.setData(sourceData, "Columns");
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### setPosition(startCell: object, endCell: object)
Positions the chart relative to cells on the worksheet.

#### Syntax
```js
chartObject.setPosition(startCell, endCell);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|startCell|object|The start cell. This is where the chart will be moved to. The start cell is the top-left or top-right cell, depending on the user's right-to-left display settings.|
|endCell|object|Optional. (Optional) The end cell. If specified, the chart's width and height will be set to fully cover up this cell/range.|

#### Returns
void

#### Examples


```js
Excel.run(function (ctx) { 
	var sheetName = "Charts";
	var rangeSelection = "A1:B4";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeSelection);
	var sourceData = sheetName + "!" + "A1:B4";
	var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("pie", range, "auto");
	chart.width = 500;
	chart.height = 300;
	chart.setPosition("C2", null);
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### Property access examples

Get a chart named "Chart1"

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.load('name');
	return ctx.sync().then(function() {
			console.log(chart.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Update a chart including renaming, positioning and resizing.

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.name="New Name";
	chart.top = 100;
	chart.left = 100;
	chart.height = 200;
	chart.width = 200;
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Rename the chart to new name, resize the chart to 200 points in both height and weight. Move Chart1 to 100 points to the top and left. 

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
	chart.name="New Name";	
	chart.top = 100;
	chart.left = 100;
	chart.height =200;
	chart.width =200;
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

