# WorksheetProtectionOptions Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents the options in sheet protection.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|allowAutoFilter|bool|Represents the worksheet protection option of allowing using auto filter feature.|1.2||
|allowDeleteColumns|bool|Represents the worksheet protection option of allowing deleting columns.|1.2||
|allowDeleteRows|bool|Represents the worksheet protection option of allowing deleting rows.|1.2||
|allowFormatCells|bool|Represents the worksheet protection option of allowing formatting cells.|1.2||
|allowFormatColumns|bool|Represents the worksheet protection option of allowing formatting columns.|1.2||
|allowFormatRows|bool|Represents the worksheet protection option of allowing formatting rows.|1.2||
|allowInsertColumns|bool|Represents the worksheet protection option of allowing inserting columns.|1.2||
|allowInsertHyperlinks|bool|Represents the worksheet protection option of allowing inserting hyperlinks.|1.2||
|allowInsertRows|bool|Represents the worksheet protection option of allowing inserting rows.|1.2||
|allowPivotTables|bool|Represents the worksheet protection option of allowing using PivotTable feature.|1.2||
|allowSort|bool|Represents the worksheet protection option of allowing using sort feature.|1.2||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

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
