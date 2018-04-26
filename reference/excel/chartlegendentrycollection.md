# ChartLegendEntryCollection Object (JavaScript API for Excel)

Represents a collection of legendEntries.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[ChartLegendEntry[]](chartlegendentry.md)|A collection of chartLegendEntry objects. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Returns the number of legendEntry in the collection.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[ChartLegendEntry](chartlegendentry.md)|Returns a legendEntry at the given index.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCount()
Returns the number of legendEntry in the collection.

#### Syntax
```js
chartLegendEntryCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItemAt(index: number)
Returns a legendEntry at the given index.

#### Syntax
```js
chartLegendEntryCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index of the legendEntry to be retrieved.|

#### Returns
[ChartLegendEntry](chartlegendentry.md)
