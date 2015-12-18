# ChartSeriesCollection Object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

Represents a collection of chart series.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|count|int|Returns the number of series in the collection. Read-only.|
|items|[ChartSeries[]](chartseries.md)|A collection of chartSeries objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getItemAt(index: number)](#getitematindex-number)|[ChartSeries](chartseries.md)|Retrieves a series based on its position in the collection|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getItemAt(index: number)
Retrieves a series based on its position in the collection

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
