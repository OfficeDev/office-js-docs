# PivotApiInvokeHelper Object (JavaScript API for Excel)

The invoke helper for Pivot APIs.

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[pivotCacheCollectionAdd(pivotCaches: PivotCacheCollection, SourceType: string, address: Range)](#pivotcachecollectionaddpivotcaches-pivotcachecollection-sourcetype-string-address-range)|[PivotCache](pivotcache.md)|Add a pivot cache|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[pivotFieldGetSubtotals(pivotField: PivotField)](#pivotfieldgetsubtotalspivotfield-pivotfield)|[Subtotals](subtotals.md)|Get pivot field sub totals|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[pivotFieldSetSubtotals(pivotField: PivotField, subtotals: Subtotals)](#pivotfieldsetsubtotalspivotfield-pivotfield-subtotals-subtotals)|void|Set pivot field sub totals|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[pivotTableAddChart(pivotTable: PivotTableExt, chartType: string, seriesBy: string)](#pivottableaddchartpivottable-pivottableext-charttype-string-seriesby-string)|[Chart](chart.md)|Add a pivot chart|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[pivotTableCollectionAdd(pivotTables: PivotTableCollectionExt, name: string, address: Range, pivotCache: PivotCache)](#pivottablecollectionaddpivottables-pivottablecollectionext-name-string-address-range-pivotcache-pivotcache)|[PivotTable](pivottable.md)|Add a pivot table|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### pivotCacheCollectionAdd(pivotCaches: PivotCacheCollection, SourceType: string, address: Range)
Add a pivot cache

#### Syntax
```js
pivotApiInvokeHelperObject.pivotCacheCollectionAdd(pivotCaches, SourceType, address);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|pivotCaches|PivotCacheCollection||
|SourceType|string||
|address|Range||

#### Returns
[PivotCache](pivotcache.md)

### pivotFieldGetSubtotals(pivotField: PivotField)
Get pivot field sub totals

#### Syntax
```js
pivotApiInvokeHelperObject.pivotFieldGetSubtotals(pivotField);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|pivotField|PivotField||

#### Returns
[Subtotals](subtotals.md)

### pivotFieldSetSubtotals(pivotField: PivotField, subtotals: Subtotals)
Set pivot field sub totals

#### Syntax
```js
pivotApiInvokeHelperObject.pivotFieldSetSubtotals(pivotField, subtotals);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|pivotField|PivotField||
|subtotals|Subtotals||

#### Returns
void

### pivotTableAddChart(pivotTable: PivotTableExt, chartType: string, seriesBy: string)
Add a pivot chart

#### Syntax
```js
pivotApiInvokeHelperObject.pivotTableAddChart(pivotTable, chartType, seriesBy);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|pivotTable|PivotTableExt||
|chartType|string|  Possible values are: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc.|
|seriesBy|string|Optional.   Possible values are: Auto, Columns, Rows|

#### Returns
[Chart](chart.md)

### pivotTableCollectionAdd(pivotTables: PivotTableCollectionExt, name: string, address: Range, pivotCache: PivotCache)
Add a pivot table

#### Syntax
```js
pivotApiInvokeHelperObject.pivotTableCollectionAdd(pivotTables, name, address, pivotCache);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|pivotTables|PivotTableCollectionExt||
|name|string||
|address|Range||
|pivotCache|PivotCache||

#### Returns
[PivotTable](pivottable.md)
