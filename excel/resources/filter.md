# Filter Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Manages the filtering of a table's column.

## Properties

None

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|criteria|[FilterCriteria](filtercriteria.md)|The currently applied filter on the given column. Read-only.|1.2||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[apply(criteria: FilterCriteria)](#applycriteria-filtercriteria)|void|Apply the given filter criteria on the given column.|1.2|
|[applyBottomItemsFilter(count: number)](#applybottomitemsfiltercount-number)|void|Apply a "Bottom Item" filter to the column for the given number of elements.|1.2|
|[applyBottomPercentFilter(percent: number)](#applybottompercentfilterpercent-number)|void|Apply a "Bottom Percent" filter to the column for the given percentage of elements.|1.2|
|[applyCellColorFilter(color: string)](#applycellcolorfiltercolor-string)|void|Apply a "Cell Color" filter to the column for the given color.|1.2|
|[applyCustomFilter(criteria1: string, criteria2: string, oper: FilterOperator)](#applycustomfiltercriteria1-string-criteria2-string-oper-filteroperator)|void|Apply a "Icon" filter to the column for the given criteria strings.|1.2|
|[applyDynamicFilter(criteria: DynamicFilterCriteria)](#applydynamicfiltercriteria-dynamicfiltercriteria)|void|Apply a "Dynamic" filter to the column.|1.2|
|[applyFontColorFilter(color: string)](#applyfontcolorfiltercolor-string)|void|Apply a "Font Color" filter to the column for the given color.|1.2|
|[applyIconFilter(icon: Icon)](#applyiconfiltericon-icon)|void|Apply a "Icon" filter to the column for the given icon.|1.2|
|[applyTopItemsFilter(count: number)](#applytopitemsfiltercount-number)|void|Apply a "Top Item" filter to the column for the given number of elements.|1.2|
|[applyTopPercentFilter(percent: number)](#applytoppercentfilterpercent-number)|void|Apply a "Top Percent" filter to the column for the given percentage of elements.|1.2|
|[applyValuesFilter(values: ()[])](#applyvaluesfiltervalues-)|void|Apply a "Values" filter to the column for the given values.|1.2|
|[clear()](#clear)|void|Clear the filter on the given column.|1.2|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### apply(criteria: FilterCriteria)
Apply the given filter criteria on the given column.

#### Syntax
```js
filterObject.apply(criteria);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
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
|:---------------|:--------|:----------|:---|
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
|:---------------|:--------|:----------|:---|
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
|:---------------|:--------|:----------|:---|
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
|:---------------|:--------|:----------|:---|
|criteria1|string|The first criteria string.|
|criteria2|string|Optional. The second criteria string.|
|oper|FilterOperator|Optional. The operator that describes how the two criteria are joined.|

#### Returns
void

### applyDynamicFilter(criteria: DynamicFilterCriteria)
Apply a "Dynamic" filter to the column.

#### Syntax
```js
filterObject.applyDynamicFilter(criteria);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|criteria|DynamicFilterCriteria|The dynamic criteria to apply.|

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
|:---------------|:--------|:----------|:---|
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
|:---------------|:--------|:----------|:---|
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
|:---------------|:--------|:----------|:---|
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
|:---------------|:--------|:----------|:---|
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
|:---------------|:--------|:----------|:---|
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
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
