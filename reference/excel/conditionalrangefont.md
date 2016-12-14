# ConditionalRangeFont Object (JavaScript API for Excel)

This object represents the font attributes (font style,, color, etc.) for an object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|bold|bool|Represents the bold status of font.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|HTML color code representation of the text color. E.g. #FF0000 represents Red.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|italic|bool|Represents the italic status of the font.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|strikethrough|bool|Represents the strikethrough status of the font.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|underline|string|Type of underline applied to the font.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|Resets the font formats.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### clear()
Resets the font formats.

#### Syntax
```js
conditionalRangeFontObject.clear();
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
