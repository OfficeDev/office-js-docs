# AutoFilter Object (JavaScript API for Excel)

Represents the AutoFilter object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|enabled|bool|Indicates if the AutoFilter is enabled or not. Read-Only. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|isDataFiltered|bool|Indicates if the AutoFilter has filter criteria. Read-Only. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|criteria|[FilterCriteria](filtercriteria.md)|Array that holds all filter criterias in an autofiltered range. Read-Only. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[apply(range: Range or string, columnIndex: number, criteria: FilterCriteria)](#applyrange-range-or-string-columnindex-number-criteria-filtercriteria)|void|Applies AutoFilter on a range and filters the column if column index and filter criteria are specified.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[clearCriteria()](#clearcriteria)|void|Clears the criteria if AutoFilter has filters|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|Returns the Range object that represents the range to which the AutoFilter applies.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getRangeOrNullObject()](#getrangeornullobject)|[Range](range.md)|If there is Range object associated with the AutoFilter, this method returns it.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[reapply()](#reapply)|void|Applies the specified Autofilter object currently on the range.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[remove()](#remove)|void|Removes the AutoFilter for the range.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### apply(range: Range or string, columnIndex: number, criteria: FilterCriteria)
Applies AutoFilter on a range and filters the column if column index and filter criteria are specified.

#### Syntax
```js
autoFilterObject.apply(range, columnIndex, criteria);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|range|Range or string|The range where the AutoFilter will apply on.|
|columnIndex|number|Optional. The column index which the AutoFilter will apply on, start from 0.|
|criteria|FilterCriteria|Optional. The filter criteria.|

#### Returns
void

### clearCriteria()
Clears the criteria if AutoFilter has filters

#### Syntax
```js
autoFilterObject.clearCriteria();
```

#### Parameters
None

#### Returns
void

### getRange()
Returns the Range object that represents the range to which the AutoFilter applies.

#### Syntax
```js
autoFilterObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getRangeOrNullObject()
If there is Range object associated with the AutoFilter, this method returns it.

#### Syntax
```js
autoFilterObject.getRangeOrNullObject();
```

#### Parameters
None

#### Returns
[Range](range.md)

### reapply()
Applies the specified Autofilter object currently on the range.

#### Syntax
```js
autoFilterObject.reapply();
```

#### Parameters
None

#### Returns
void

### remove()
Removes the AutoFilter for the range.

#### Syntax
```js
autoFilterObject.remove();
```

#### Parameters
None

#### Returns
void
