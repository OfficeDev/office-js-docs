# BindingSelectionChangedEventArgs Object (JavaScript API for Excel)

Provides information about the binding that raised the SelectionChanged event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|columnCount|int|Gets the number of columns selected.|[1.2, 1.3](../excel-requirement.md)|
|rowCount|int|Gets the number of rows selected.|[1.2, 1.3](../excel-requirement.md)|
|startColumn|int|Gets the index of the first column of the selection (zero-based).|[1.2, 1.3](../excel-requirement.md)|
|startRow|int|Gets the index of the first row of the selection (zero-based).|[1.2, 1.3](../excel-requirement.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|binding|[public](public.md)|Gets the Binding object that represents the binding that raised the SelectionChanged event.|[1.2, 1.3](../excel-requirement.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../reqset/excel-requirement.md)|

## Method Details


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
