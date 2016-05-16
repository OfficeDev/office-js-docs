# RangeSort Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Manages sorting operations on Range objects.

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: SortOrientation, method: SortMethod)](#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-sortorientation-method-sortmethod)|void|Perform a sort operation.|1.2|

## Method Details


### apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: SortOrientation, method: SortMethod)
Perform a sort operation.

#### Syntax
```js
rangeSortObject.apply(fields, matchCase, hasHeaders, orientation, method);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|fields|SortField[]|The list of conditions to sort on.|
|matchCase|bool|Optional. Whether to have the casing impact string ordering.|
|hasHeaders|bool|Optional. Whether the range has a header.|
|orientation|SortOrientation|Optional. Whether the operation is sorting rows or columns.|
|method|SortMethod|Optional. The ordering method used for Chinese characters.|

#### Returns
void
