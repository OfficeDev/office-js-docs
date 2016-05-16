# TableSort Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Manages sorting operations on Table objects.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|matchCase|bool|Represents whether the casing impacted the last sort of the table. Read-only.|1.2||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|fields|[SortField](sortfield.md)|Represents the current conditions used to last sort the table. Read-only.|1.2||
|method|[SortMethod](sortmethod.md)|Represents Chinese character ordering method last used to sort the table. Read-only.|1.2||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[apply(fields: SortField[], matchCase: bool, method: SortMethod)](#applyfields-sortfield-matchcase-bool-method-sortmethod)|void|Perform a sort operation.|1.2|
|[clear()](#clear)|void|Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.|1.2|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|
|[reapply()](#reapply)|void|Reapplies the current sorting parameters to the table.|1.2|

## Method Details


### apply(fields: SortField[], matchCase: bool, method: SortMethod)
Perform a sort operation.

#### Syntax
```js
tableSortObject.apply(fields, matchCase, method);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|fields|SortField[]|The list of conditions to sort on.|
|matchCase|bool|Optional. Whether to have the casing impact string ordering.|
|method|SortMethod|Optional. The ordering method used for Chinese characters.|

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
|:---------------|:--------|:----------|:---|
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
