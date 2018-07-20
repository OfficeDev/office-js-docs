# PageBreak Object (JavaScript API for Excel)

Represents the options in page layout margins.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|columnIndex|int|Represents the column index for the page break Read-only.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|rowIndex|int|Represents the row index for the page break Read-only.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes a page break object.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|[getStartCell()](#getstartcell)|[Range](range.md)|Gets the first cell after the page break.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### delete()
Deletes a page break object.

#### Syntax
```js
pageBreakObject.delete();
```

#### Parameters
None

#### Returns
void

### getStartCell()
Gets the first cell after the page break.

#### Syntax
```js
pageBreakObject.getStartCell();
```

#### Parameters
None

#### Returns
[Range](range.md)
