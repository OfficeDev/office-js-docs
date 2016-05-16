# Binding Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents an Office.js binding that is defined in the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|Represents binding identifier. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|type|string|Returns the type of the binding. Read-only.|1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getRange()](#getrange)|[Range](range.md)|Returns the range represented by the binding. Will throw an error if binding is not of the correct type.|1.1|
|[getTable()](#gettable)|[Table](table.md)|Returns the table represented by the binding. Will throw an error if binding is not of the correct type.|1.1|
|[getText()](#gettext)|string|Returns the text represented by the binding. Will throw an error if binding is not of the correct type.|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### getRange()
Returns the range represented by the binding. Will throw an error if binding is not of the correct type.

#### Syntax
```js
bindingObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getTable()
Returns the table represented by the binding. Will throw an error if binding is not of the correct type.

#### Syntax
```js
bindingObject.getTable();
```

#### Parameters
None

#### Returns
[Table](table.md)

### getText()
Returns the text represented by the binding. Will throw an error if binding is not of the correct type.

#### Syntax
```js
bindingObject.getText();
```

#### Parameters
None

#### Returns
string

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
