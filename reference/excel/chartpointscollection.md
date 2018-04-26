# ChartPointsCollection Object (JavaScript API for Excel)

A collection of all the chart points within a series inside a chart.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|count|int|Returns the number of chart points in the series. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[ChartPoint[]](chartpoint.md)|A collection of chartPoints objects. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Returns the number of chart points in the series.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getFirst()](#getfirst)|[ChartPoint](chartpoint.md)|Gets the first point in the series.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[ChartPoint](chartpoint.md)|Retrieve a point based on its position within the series.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getLast()](#getlast)|[ChartPoint](chartpoint.md)|Gets the last point in the series.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCount()
Returns the number of chart points in the series.

#### Syntax
```js
chartPointsCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getFirst()
Gets the first point in the series.

#### Syntax
```js
chartPointsCollectionObject.getFirst();
```

#### Parameters
None

#### Returns
[ChartPoint](chartpoint.md)

### getItemAt(index: number)
Retrieve a point based on its position within the series.

#### Syntax
```js
chartPointsCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[ChartPoint](chartpoint.md)

#### Examples
Set the border color for the first points in the points collection

```js
Excel.run(function (ctx) { 
	var points = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
	points.getItemAt(0).format.fill.setSolidColor("8FBC8F");
	return ctx.sync().then(function() {
		console.log("Point Border Color Changed");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### getLast()
Gets the last point in the series.

#### Syntax
```js
chartPointsCollectionObject.getLast();
```

#### Parameters
None

#### Returns
[ChartPoint](chartpoint.md)
### Property access examples

Get the names of points in the points collection

```js
Excel.run(function (ctx) { 
	var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
	pointsCollection.load('items');
	return ctx.sync().then(function() {
		console.log("Points Collection loaded");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Get the number of points

```js
Excel.run(function (ctx) { 
	var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
	pointsCollection.load('count');
	return ctx.sync().then(function() {
		console.log("points: Count= " + pointsCollection.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
