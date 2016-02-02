# Filter Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

Manages the filtering of a table's column.

## Properties

None

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|criteria|[FilterCriteria](filtercriteria.md)|The currently applied filter on the given column. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[apply(criteria: FilterCriteria)](#applycriteria-filtercriteria)|void|Apply the given filter criteria on the given column.|
|[applyBottomItemsFilter(count: number)](#applybottomitemsfiltercount-number)|void|Apply a "Bottom Item" filter to the column for the given number of elements.|
|[applyBottomPercentFilter(percent: number)](#applybottompercentfilterpercent-number)|void|Apply a "Bottom Percent" filter to the column for the given percentage of elements.|
|[applyCellColorFilter(color: string)](#applycellcolorfiltercolor-string)|void|Apply a "Cell Color" filter to the column for the given color.|
|[applyCustomFilter(criteria1: string, criteria2: string, oper: FilterOperator)](#applycustomfiltercriteria1-string-criteria2-string-oper-filteroperator)|void|Apply a "Icon" filter to the column for the given criteria strings.|
|[applyDynamicFilter(criteria: string)](#applydynamicfiltercriteria-string)|void|Apply a "Dynamic" filter to the column.|
|[applyFontColorFilter(color: string)](#applyfontcolorfiltercolor-string)|void|Apply a "Font Color" filter to the column for the given color.|
|[applyIconFilter(icon: Icon)](#applyiconfiltericon-icon)|void|Apply a "Icon" filter to the column for the given icon.|
|[applyTopItemsFilter(count: number)](#applytopitemsfiltercount-number)|void|Apply a "Top Item" filter to the column for the given number of elements.|
|[applyTopPercentFilter(percent: number)](#applytoppercentfilterpercent-number)|void|Apply a "Top Percent" filter to the column for the given percentage of elements.|
|[applyValuesFilter(values: ()[])](#applyvaluesfiltervalues-)|void|Apply a "Values" filter to the column for the given values.|
|[clear()](#clear)|void|Clear the filter on the given column.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### apply(criteria: FilterCriteria)
Apply the given filter criteria on the given column.

#### Syntax
```js
filterObject.apply(criteria);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|criteria|FilterCriteria|The criteria to apply.|

#### Returns
void

### applyBottomItemsFilter(count: number)
Apply a "Bottom Item" filter to the column for the given number of elements.

#### Syntax
```js
filterObject.applyBottomItemsFilter(count);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|count|number|The number of elements from the bottom to show.|

#### Returns
void

### applyBottomPercentFilter(percent: number)
Apply a "Bottom Percent" filter to the column for the given percentage of elements.

#### Syntax
```js
filterObject.applyBottomPercentFilter(percent);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|percent|number|The percentage of elements from the bottom to show.|

#### Returns
void

### applyCellColorFilter(color: string)
Apply a "Cell Color" filter to the column for the given color.

#### Syntax
```js
filterObject.applyCellColorFilter(color);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|color|string|The background color of the cells to show.|

#### Returns
void

### applyCustomFilter(criteria1: string, criteria2: string, oper: FilterOperator)
Apply a "Icon" filter to the column for the given criteria strings.

#### Syntax
```js
filterObject.applyCustomFilter(criteria1, criteria2, oper);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|criteria1|string|The first criteria string.|
|criteria2|string|Optional. The second criteria string.|
|oper|FilterOperator|Optional. The operator that describes how the two criteria are joined.|

#### Returns
void

### applyDynamicFilter(criteria: string)
Apply a "Dynamic" filter to the column.

#### Syntax
```js
filterObject.applyDynamicFilter(criteria);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|criteria|string|The dynamic criteria to apply.  Possible values are: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday|

#### Returns
void

### applyFontColorFilter(color: string)
Apply a "Font Color" filter to the column for the given color.

#### Syntax
```js
filterObject.applyFontColorFilter(color);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|color|string|The font color of the cells to show.|

#### Returns
void

### applyIconFilter(icon: Icon)
Apply a "Icon" filter to the column for the given icon.

#### Syntax
```js
filterObject.applyIconFilter(icon);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|icon|Icon|The icons of the cells to show.|

#### Returns
void

### applyTopItemsFilter(count: number)
Apply a "Top Item" filter to the column for the given number of elements.

#### Syntax
```js
filterObject.applyTopItemsFilter(count);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|count|number|The number of elements from the top to show.|

#### Returns
void

### applyTopPercentFilter(percent: number)
Apply a "Top Percent" filter to the column for the given percentage of elements.

#### Syntax
```js
filterObject.applyTopPercentFilter(percent);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|percent|number|The percentage of elements from the top to show.|

#### Returns
void

### applyValuesFilter(values: ()[])
Apply a "Values" filter to the column for the given values.

#### Syntax
```js
filterObject.applyValuesFilter(values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|values|()[]|The list of values to show.|

#### Returns
void

### clear()
Clear the filter on the given column.

#### Syntax
```js
filterObject.clear();
```

#### Parameters
None

#### Returns
void

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
