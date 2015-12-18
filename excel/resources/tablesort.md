# TableSort Object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

Manages sorting operations on Table objects.

_**Note**: This is a proposed feature that is still under design phase and hence not yet available as part of the product. The specification is being made available for community review and feedback. The final design may change. Help us make this feature better by providing your feedback [here](https://github.com/OfficeDev/office-js-docs/issues/new?title=ExcelJs-1.2-OpenSpec-sort)._

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|matchCase|bool|Represents whether the casing impacted the last sort of the table. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|fields|[SortField](sortfield.md)|Represents the current conditions used to last sort the table. Read-only.|
|method|[SortMethod](sortmethod.md)|Represents Chinese character ordering method last used to sort the table. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[apply(fields: SortField[, matchCase: [Optional, method: [Optional)](#applyfields-sortfield-matchcase-optional-method-optional)|void|Perform a sort operation.|
|[clear()](#clear)|void|Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[reapply()](#reapply)|void|Reapplies the current sorting parameters to the table.|

## Method Details


### apply(fields: SortField[, matchCase: [Optional, method: [Optional)
Perform a sort operation.

#### Syntax
```js
tableSortObject.apply(fields, matchCase, method);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|fields|SortField[|The list of conditions to sort on.|
|matchCase|[Optional|Optional. Whether to have the casing impact string ordering.|
|method|[Optional|Optional. The ordering method used for Chinese characters.|

#### Returns
void

### clear()
Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.

#### Syntax
```js
tableSortObject.clear();
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

### reapply()
Reapplies the current sorting parameters to the table.

#### Syntax
```js
tableSortObject.reapply();
```

#### Parameters
None

#### Returns
void
