# RangeSort Object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

Manages sorting operations on Range objects.

_**Note**: This is a proposed feature that is still under design phase and hence not yet available as part of the product. The specification is being made available for community review and feedback. The final design may change. Help us make this feature better by providing your feedback [here](https://github.com/OfficeDev/office-js-docs/issues/new?title=ExcelJs-1.2-OpenSpec-sort)._

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[apply(fields: SortField[, matchCase: [Optional, hasHeaders: [Optional, orientation: [Optional, method: [Optional)](#applyfields-sortfield-matchcase-optional-hasheaders-optional-orientation-optional-method-optional)|void|Perform a sort operation.|

## Method Details


### apply(fields: SortField[, matchCase: [Optional, hasHeaders: [Optional, orientation: [Optional, method: [Optional)
Perform a sort operation.

#### Syntax
```js
rangeSortObject.apply(fields, matchCase, hasHeaders, orientation, method);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|fields|SortField[|The list of conditions to sort on.|
|matchCase|[Optional|Optional. Whether to have the casing impact string ordering.|
|hasHeaders|[Optional|Optional. Whether the range has a header.|
|orientation|[Optional|Optional. Whether the operation is sorting rows or columns.|
|method|[Optional|Optional. The ordering method used for Chinese characters.|

#### Returns
void
