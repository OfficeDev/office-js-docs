# DataValidation Object (JavaScript API for Excel)

Represents the data validation applied to the current range.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|ignoreBlanks|bool|Ignore blanks: no data validation will be performed on blank cells, it defaults to true.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Type of the data validation, see Excel.DataValidationType for details. Read-only. Possible values are: Invalid, None, WholeNumber, Decimal, List, Date, Time, TextLength, Inconsistent, MixedCriteria.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|errorAlert|[DataValidationErrorAlert](datavalidationerroralert.md)|Error alert when user enters invalid data.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|prompt|[DataValidationPrompt](datavalidationprompt.md)|Prompt when users select a cell.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|rule|[DataValidationRule](datavalidationrule.md)|Data Validation rule that contains different type of data validation criteria.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|Clears the data validation from the current range.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### clear()
Clears the data validation from the current range.

#### Syntax
```js
dataValidationObject.clear();
```

#### Parameters
None

#### Returns
void
