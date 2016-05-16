# BindingCollection Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents the collection of all the binding objects that are part of the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|count|int|Returns the number of bindings in the collection. Read-only.|1.1||
|items|[Binding[]](binding.md)|A collection of binding objects. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(range: Range or string, bindingType: string, id: string)](#addrange-range-or-string-bindingtype-string-id-string)|[Binding](binding.md)|Add a new binding to a particular Range.|1.3|
|[addFromNamedItem(name: string, bindingType: string, id: string)](#addfromnameditemname-string-bindingtype-string-id-string)|[Binding](binding.md)|Add a new binding based on a named item in the workbook.|1.3|
|[addFromSelection(bindingType: string, id: string)](#addfromselectionbindingtype-string-id-string)|[Binding](binding.md)|Add a new binding based on the current selection.|1.3|
|[getItem(id: string)](#getitemid-string)|[Binding](binding.md)|Gets a binding object by ID.|1.1|
|[getItemAt(index: number)](#getitematindex-number)|[Binding](binding.md)|Gets a binding object based on its position in the items array.|1.1|
|[getItemOrNull(id: string)](#getitemornullid-string)|[Binding](binding.md)|Gets a binding object by ID. If the binding object does not exist, the return object's isNull property will be true.|1.3|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### add(range: Range or string, bindingType: string, id: string)
Add a new binding to a particular Range.

#### Syntax
```js
bindingCollectionObject.add(range, bindingType, id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|range|Range or string|Range to bind the binding to. May be an Excel Range object, or a string. If string, must contain the full address, including the sheet name|
|bindingType|string|Type of binding.|
|id|string|Name of binding.|

#### Returns
[Binding](binding.md)

### addFromNamedItem(name: string, bindingType: string, id: string)
Add a new binding based on a named item in the workbook.

#### Syntax
```js
bindingCollectionObject.addFromNamedItem(name, bindingType, id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|name|string|Name from which to create binding.|
|bindingType|string|Type of binding.|
|id|string|Name of binding.|

#### Returns
[Binding](binding.md)

### addFromSelection(bindingType: string, id: string)
Add a new binding based on the current selection.

#### Syntax
```js
bindingCollectionObject.addFromSelection(bindingType, id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|bindingType|string|Type of binding.|
|id|string|Name of binding.|

#### Returns
[Binding](binding.md)

### getItem(id: string)
Gets a binding object by ID.

#### Syntax
```js
bindingCollectionObject.getItem(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|id|string|Id of the binding object to be retrieved.|

#### Returns
[Binding](binding.md)

### getItemAt(index: number)
Gets a binding object based on its position in the items array.

#### Syntax
```js
bindingCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[Binding](binding.md)

### getItemOrNull(id: string)
Gets a binding object by ID. If the binding object does not exist, the return object's isNull property will be true.

#### Syntax
```js
bindingCollectionObject.getItemOrNull(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|id|string|Id of the binding object to be retrieved.|

#### Returns
[Binding](binding.md)

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
