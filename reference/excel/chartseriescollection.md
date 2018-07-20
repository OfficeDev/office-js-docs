# ChartSeriesCollection Object (JavaScript API for Excel)

Represents a collection of chart series.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|count|int|Returns the number of series in the collection. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[ChartSeries[]](chartseries.md)|A collection of chartSeries objects. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(name: string, index: number)](#addname-string-index-number)|[ChartSeries](chartseries.md)|Add a new series to the collection. The new added series is not visible until set valuesx axis valuesbubble sizes for it (depending on chart type).|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Returns the number of series in the collection.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getFirst()](#getfirst)|[ChartSeries](chartseries.md)|Gets the first series in the collection.|[ApiSet.InProgressFeatures.GetFirstGetLast](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[ChartSeries](chartseries.md)|Retrieves a series based on its position in the collection.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getLast()](#getlast)|[ChartSeries](chartseries.md)|Gets the last series in the collection.|[ApiSet.InProgressFeatures.GetFirstGetLast](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(name: string, index: number)
Add a new series to the collection. The new added series is not visible until set valuesx axis valuesbubble sizes for it (depending on chart type).

#### Syntax
```js
chartSeriesCollectionObject.add(name, index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Optional. Optional. Name of the series.|
|index|number|Optional. Optional. Index value of the series to be added. Zero-indexed.|

#### Returns
[ChartSeries](chartseries.md)

### getCount()
Returns the number of series in the collection.

#### Syntax
```js
chartSeriesCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getFirst()
Gets the first series in the collection.

#### Syntax
```js
chartSeriesCollectionObject.getFirst();
```

#### Parameters
None

#### Returns
[ChartSeries](chartseries.md)

### getItemAt(index: number)
Retrieves a series based on its position in the collection.

#### Syntax
```js
chartSeriesCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[ChartSeries](chartseries.md)

#### Examples

Get the name of the first series in the series collection.

```js
Excel.run(function (ctx) { 
	var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
	seriesCollection.load('items');
	return ctx.sync().then(function() {
		console.log(seriesCollection.items[0].name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getLast()
Gets the last series in the collection.

#### Syntax
```js
chartSeriesCollectionObject.getLast();
```

#### Parameters
None

#### Returns
[ChartSeries](chartseries.md)
### Property access examples
Getting the names of series in the series collection.

```js
Excel.run(function (ctx) { 
	var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
	seriesCollection.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < seriesCollection.items.length; i++)
		{
			console.log(seriesCollection.items[i].name);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Get the number of chart series in collection.

```js
Excel.run(function (ctx) { 
	var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
	seriesCollection.load('count');
	return ctx.sync().then(function() {
		console.log("series: Count= " + seriesCollection.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

