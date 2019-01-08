# BasicDataValidation Object (JavaScript API for Excel)

Represents the Basic Type data validation criteria.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|formula1|object|Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell). With the ternary operators Between and NotBetween, specifies the lower bound operand.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|formula2|object|With the ternary operators Between and NotBetween, specifies the upper bound operand. Is not used with the binary operators, such as GreaterThan.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|operator|[DataValidationOperator](datavalidationoperator.md)|The operator to use for validating the data.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

