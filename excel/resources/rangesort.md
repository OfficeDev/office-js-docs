# RangeSort Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

Manages sorting operations on Range objects.

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)](#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|void|Perform a sort operation.|

## Method Details


### apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)
Perform a sort operation.

#### Syntax
```js
rangeSortObject.apply(fields, matchCase, hasHeaders, orientation, method);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|fields|SortField[]|The list of conditions to sort on.|
|matchCase|bool|Optional. Whether to have the casing impact string ordering.|
|hasHeaders|bool|Optional. Whether the range has a header.|
|orientation|string|Optional. Whether the operation is sorting rows or columns.  Possible values are: Rows, Columns|
|method|string|Optional. The ordering method used for Chinese characters.  Possible values are: PinYin, StrokeCount|

#### Returns
void
