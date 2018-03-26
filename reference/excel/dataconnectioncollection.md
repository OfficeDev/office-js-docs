# DataConnectionCollection Object (JavaScript API for Excel)

Represents a collection of all the Data Connections that are part of the workbook or worksheet.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[DataConnection[]](dataconnection.md)|A collection of dataConnection objects. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[refreshAll()](#refreshall)|void|Refreshes all the Data Connections in the collection.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### refreshAll()
Refreshes all the Data Connections in the collection.

#### Syntax
```js
dataConnectionCollectionObject.refreshAll();
```

#### Parameters
None

#### Returns
void
