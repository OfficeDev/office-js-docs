# ChartSeries Object (JavaScript API for Excel)

Represents a series in a chart.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|axisGroup|string|Returns or sets the group for the specified series. ReadWrite Possible values are: Primary, Secondary.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|bubbleScale|double|Returns or sets the scale factor for bubbles in the specified chart group. Can be an integer value from 0 (zero) to 300, corresponding to a percentage of the default size. Applies only to bubble charts. ReadWrite.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|chartType|string|Represents the chart type of a series. Possible values are: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc..|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|doughnutHoleSize|double|Represents the doughnut hole size of a chart series.  Only valid on doughnut and doughnutExploded charts.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|explosion|int|Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). ReadWrite.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|filtered|bool|Boolean value representing if the series is filtered or not. Not applicable for surface charts.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|firstSliceAngle|int|Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360. ReadWrite|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|gapWidth|double|Represents the gap width of a chart series.  Only valid on bar and column charts, as well as|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|hasDataLabels|bool|Boolean value representing if the series has data labels or not.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|hasLeaderLines|bool|True if Microsoft Excel show leaderlines for each datalabel in series. ReadWrite.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|invertIfNegative|bool|True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. ReadWrite.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|markerBackgroundColor|string|Represents markers background color of a chart series.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|markerForegroundColor|string|Represents markers foreground color of a chart series.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|markerSize|int|Represents marker size of a chart series.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|markerStyle|string|Represents marker style of a chart series. Possible values are: Automatic, None, Square, Diamond, Triangle, X, Star, Dot, Dash, Circle, Plus, Picture.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Represents the name of a series in a chart.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|overlap|int|Specifies how bars and columns are positioned. Can be a value between -100 and 100. Applies only to 2-D bar and 2-D column charts. ReadWrite.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|plotOrder|int|Represents the plot order of a chart series within the chart group.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|secondPlotSize|int|Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. ReadWrite.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|showShadow|bool|Boolean value representing if the series has shadow or not.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|smooth|bool|Boolean value representing if the series is smooth or not. Only for line and scatter charts.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|splitType|string|Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split. ReadWrite. Possible values are: SplitByPosition, SplitByValue, SplitByPercentValue, SplitByCustomSplit.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|splitValue|double|Returns or sets the threshold value separating the two sections of either a pie of pie chart or a bar of pie chart. ReadWrite.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Returns the value that represents the series type. Read-only. Possible values are: Column, Bar, Bar3D, Line, Pie, XYScatter, Area, Area3D, Doughnut, Radar, Surface3D, Column3D.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|varyByCategories|bool|True if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series. ReadWrite.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|dataLabels|[ChartDataLabels](chartdatalabels.md)|Represents a collection of all dataLabels in the series. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|format|[ChartSeriesFormat](chartseriesformat.md)|Represents the formatting of a chart series, which includes fill and line formatting. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|points|[ChartPointsCollection](chartpointscollection.md)|Represents a collection of all points in the series. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|trendlines|[ChartTrendlineCollection](charttrendlinecollection.md)|Represents a collection of trendlines in the series. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|xErrorBars|[ChartErrorBars](charterrorbars.md)|Represents the error bar object for a chart series. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|yErrorBars|[ChartErrorBars](charterrorbars.md)|Represents the error bar object for a chart series. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the chart series.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[setBubbleSizes(sourceData: Range)](#setbubblesizessourcedata-range)|void|Set bubble sizes for a chart series. Only works for bubble charts.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[setValues(sourceData: Range)](#setvaluessourcedata-range)|void|Set values for a chart series. For scatter chart, it means Y axis values.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[setXAxisValues(sourceData: Range)](#setxaxisvaluessourcedata-range)|void|Set values of X axis for a chart series. Only works for scatter charts.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### delete()
Deletes the chart series.

#### Syntax
```js
chartSeriesObject.delete();
```

#### Parameters
None

#### Returns
void

### setBubbleSizes(sourceData: Range)
Set bubble sizes for a chart series. Only works for bubble charts.

#### Syntax
```js
chartSeriesObject.setBubbleSizes(sourceData);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|sourceData|Range|The Range object corresponding to the source data.|

#### Returns
void

### setValues(sourceData: Range)
Set values for a chart series. For scatter chart, it means Y axis values.

#### Syntax
```js
chartSeriesObject.setValues(sourceData);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|sourceData|Range|The Range object corresponding to the source data.|

#### Returns
void

### setXAxisValues(sourceData: Range)
Set values of X axis for a chart series. Only works for scatter charts.

#### Syntax
```js
chartSeriesObject.setXAxisValues(sourceData);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|sourceData|Range|The Range object corresponding to the source data.|

#### Returns
void
### Property access examples

Rename the 1st series of Chart1 to "New Series Name"

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.series.getItemAt(0).name = "New Series Name";
	return ctx.sync().then(function() {
			console.log("Series1 Renamed");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
