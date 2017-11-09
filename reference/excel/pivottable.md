# PivotTable Object (JavaScript API for Excel)

Represents an Excel PivotTable.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|columnGrandTotals|bool|True if the PivotTable report shows grand totals for columns.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|Id of the PivotTable. Read-only.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Name of the PivotTable.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|rowGrandTotals|bool|True if the PivotTable report shows grand totals for rows.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|worksheet|[Worksheet](worksheet.md)|The worksheet containing the current PivotTable. Read-only.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[refresh()](#refresh)|void|Refreshes the PivotTable.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[rowAxisLayout(RowLayout: string)](#rowaxislayoutrowlayout-string)|void|This method is used for simultaneously setting layout options for all existing PivotFields.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|[subtotalLocation(Location: string)](#subtotallocationlocation-string)|void|This method changes the subtotal location for all existing PivotFields. Changing the subtotal location has an immediate visual effect only for fields in outline form, but it will be set for fields in tabular form as well.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### refresh()
Refreshes the PivotTable.

#### Syntax
```js
pivotTableObject.refresh();
```

#### Parameters
None

#### Returns
void

### rowAxisLayout(RowLayout: string)
This method is used for simultaneously setting layout options for all existing PivotFields.

#### Syntax
```js
pivotTableObject.rowAxisLayout(RowLayout);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|RowLayout|string|Required LayoutRowType.  Possible values are: CompactRow, TabularRow, OutlineRow|

#### Returns
void

### subtotalLocation(Location: string)
This method changes the subtotal location for all existing PivotFields. Changing the subtotal location has an immediate visual effect only for fields in outline form, but it will be set for fields in tabular form as well.

#### Syntax
```js
pivotTableObject.subtotalLocation(Location);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|Location|string|Required SubtotalLocationType.  Possible values are: AtTop, AtBottom|

#### Returns
void
