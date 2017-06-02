# ChartSeries Object (JavaScript API for Excel)

Represents a series in a chart.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|name|string|Represents the name of a series in a chart.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartSeriesFormat](chartseriesformat.md)|Represents the formatting of a chart series, which includes fill and line formatting. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|points|[ChartPointsCollection](chartpointscollection.md)|Represents a collection of all points in the series. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|trendlines|[ChartTrendlineCollection](charttrendlinecollection.md)|Represents a collection of Trendlines in the series. Read-only.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the chart series.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|
|[setBubbleSizes(sourceData: Range)](#setbubblesizessourcedata-range)|void|Set bubble sizes for a chart series. Only works for bubble charts.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|
|[setValues(sourceData: Range)](#setvaluessourcedata-range)|void|Set values for a chart series. For scatter chart, it means Y axis values.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|
|[setXAxisValues(sourceData: Range)](#setxaxisvaluessourcedata-range)|void|Set values of X axis for a chart series. Only works for scatter charts.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|

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
