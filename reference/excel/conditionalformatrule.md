# ConditionalFormatRule Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents a rule, for all traditional ruleformat pairings.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|blank|bool|If true, formats blank cells. If false, formats not blank cells.|1.3||
|datesOccurring|string|Takes a date-like enumeration to format cells based on. Possible values are: Yesterday, Today, Tomorrow, LastSevenDays, LastWeek, ThisWeek, NextWeek, LastMonth, ThisMonth, NextMonth.|1.3||
|error|bool|If true, formats cells with errors. If false, formats cells without errors.|1.3||
|formula|string|Represents all formula-based expressions, as a string.|1.3||
|type|string|Tells what type of rule is being applied. If multiple exist, the first occurred is the one that is applied. Possible values are: Formula, Between, NotBetween, Count, Percent, Average, Blank, Unique, Error, TextContains, DateOccurring.|1.3||
|unique|bool|If true, formats unique cells. If false, formats duplicate cells.|1.3||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|average|[ConditionalFormatAverageRule](conditionalformataveragerule.md)|Represents all average-based rules for conditional formatting.|1.3||
|between|[ConditionalFormatBetweenRule](conditionalformatbetweenrule.md)|Represents all rules that have a between formula of items to be formatted.|1.3||
|count|[ConditionalFormatCountValueRule](conditionalformatcountvaluerule.md)|Represents the top or bottom count of items to be formatted.|1.3||
|notBetween|[ConditionalFormatBetweenRule](conditionalformatbetweenrule.md)|Represents all rules that have a not between formula of items to be formatted.|1.3||
|numberRule|[ConditionalFormatNumericalRule](conditionalformatnumericalrule.md)|Represents a numerical comparison rule type.|1.3||
|percent|[ConditionalFormatPercentValueRule](conditionalformatpercentvaluerule.md)|Represents the top or bottom percentage of items to be formatted.|1.3||
|textContains|[ConditionalFormatStringRule](conditionalformatstringrule.md)|Represents all text-contains or string based rules on cells.|1.3||

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
