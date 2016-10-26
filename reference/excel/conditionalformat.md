# ConditionalFormat Object (JavaScript API for Excel)

An object encapsulating a conditional format's range, format, rule, and other properties.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|priority|int|The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also|[1.4](../excel-requirement.md)|
|stopIfTrue|bool|If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.|[1.4](../excel-requirement.md)|
|type|string|A type of conditional format. Only one can be set at a time. Read-Only. Read-only. Possible values are: Custom, DataBar, ColorScale, IconSet.|[1.4](../excel-requirement.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|dataBarOrNull|[ConditionalFormatDataBar](conditionalformatdatabar.md)|Represents databars with customizable color, gradient, axis, and range format options. Read-only.|[1.4](../excel-requirement.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes this conditional format.|[1.4](../reqset/excel-requirement.md)|
|[getRangeOrNull()](#getrangeornull)|[Range](range.md)|Returns the range the conditonal format is applied to or null if the range is discontiguous. Read-only.|[1.4](../reqset/excel-requirement.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../reqset/excel-requirement.md)|

## Method Details


### delete()
Deletes this conditional format.

#### Syntax
```js
conditionalFormatObject.delete();
```

#### Parameters
None

#### Returns
void

### getRangeOrNull()
Returns the range the conditonal format is applied to or null if the range is discontiguous. Read-only.

#### Syntax
```js
conditionalFormatObject.getRangeOrNull();
```

#### Parameters
None

#### Returns
[Range](range.md)

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
### Property access examples
