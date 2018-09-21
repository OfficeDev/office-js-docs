# PivotLayout Object (JavaScript API for Excel)

Represents the visual layout of the PivotTable.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|enableFieldList|bool|True if the field list should be shown or hidden from the UI.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|showColumnGrandTotals|bool|True if the PivotTable report shows grand totals for columns.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|showRowGrandTotals|bool|True if the PivotTable report shows grand totals for rows.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|subtotalLocation|string|This property indicates the SubtotalLocationType of all fields on the PivotTable. If fields have different states, this will be null. Possible values are: AtTop, AtBottom.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|layoutType|[PivotLayoutType](pivotlayouttype.md)|This property indicates the PivotLayoutType of all fields on the PivotTable. If fields have different states, this will be null.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getColumnLabelRange()](#getcolumnlabelrange)|[Range](range.md)|Returns the range where the PivotTable's column labels reside.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Returns the range where the PivotTable's data values reside.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getFilterAxisRange()](#getfilteraxisrange)|[Range](range.md)|Returns the range of the PivotTable's filter area.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|Returns the range the PivotTable exists on, excluding the filter area.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowLabelRange()](#getrowlabelrange)|[Range](range.md)|Returns the range where the PivotTable's row labels reside.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getColumnLabelRange()
Returns the range where the PivotTable's column labels reside.

#### Syntax
```js
pivotLayoutObject.getColumnLabelRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getDataBodyRange()
Returns the range where the PivotTable's data values reside.

#### Syntax
```js
pivotLayoutObject.getDataBodyRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getFilterAxisRange()
Returns the range of the PivotTable's filter area.

#### Syntax
```js
pivotLayoutObject.getFilterAxisRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getRange()
Returns the range the PivotTable exists on, excluding the filter area.

#### Syntax
```js
pivotLayoutObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getRowLabelRange()
Returns the range where the PivotTable's row labels reside.

#### Syntax
```js
pivotLayoutObject.getRowLabelRange();
```

#### Parameters
None

#### Returns
[Range](range.md)
