# ConditionalIconCriterion Object (JavaScript API for Excel)

Represents an Icon Criterion which contains a type, value, an Operator, and an optional custom icon, if not using an iconset.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|formula|object|A number or a formula depending on the type.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|customIcon|[Icon](icon.md)|The custom icon for the current criterion if different from the default IconSet, else null will be returned.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|operator|[ConditionalIconCriterionOperator](conditionaliconcriterionoperator.md)|GreaterThan or GreaterThanOrEqual for each of the rule type for the Icon conditional format.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|type|[ConditionalFormatIconRuleType](conditionalformaticonruletype.md)|What the icon conditional formula should be based on.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

