# ConditionalFormat Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents an Excel conditional format.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|priority|int|The priority (or index) within the conditional format collection of the current range.|1.3||
|reverse|bool|Returns true if the conditional format does the reverse of its current settings.|1.3||
|stopIfTrue|bool|If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.|1.3||
|type|string|The type of the conditional format. Read-Only. Read-only. Possible values are: Custom, DataBar, ColorScale, IconSet.|1.3||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|colorScaleOrNull|[ConditionalFormatColorScale](conditionalformatcolorscale.md)|Represents a conditional format that applies a color scale on a range based on minimum, maximum, and Read-only.|1.3||
|customOrNull|[ConditionalFormatCustom](conditionalformatcustom.md)|A custom conditional format and rule. Read-only.|1.3||
|dataBarOrNull|[ConditionalFormatDataBar](conditionalformatdatabar.md)|Represents databars with customizable color, gradient, axis, and range format options. Read-only.|1.3||
|iconOrNull|[ConditionalFormatIcon](conditionalformaticon.md)|Represents a conditional format that applies icons based on criteria. The criteria can Read-only.|1.3||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes this conditional format from all ranges it affects.|1.3|
|[deleteFromCurrentRange()](#deletefromcurrentrange)|void|Removes this conditional format from the current range. The conditional format will only be removed for the cells of the|1.3|
|[getRangeOrNull()](#getrangeornull)|[Range](range.md)|Gets the entire range the conditional format affects, unless it is discontiguous, in which case it will return null.|1.3|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### delete()
Deletes this conditional format from all ranges it affects.

#### Syntax
```js
conditionalFormatObject.delete();
```

#### Parameters
None

#### Returns
void

### deleteFromCurrentRange()
Removes this conditional format from the current range. The conditional format will only be removed for the cells of the

#### Syntax
```js
conditionalFormatObject.deleteFromCurrentRange();
```

#### Parameters
None

#### Returns
void

### getRangeOrNull()
Gets the entire range the conditional format affects, unless it is discontiguous, in which case it will return null.

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
