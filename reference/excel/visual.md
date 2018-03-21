# Visual Object (JavaScript API for Excel)

Represents a visual object in a workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|The unique id of visual, not the guid of VisualDefinition. Read-only.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Delete the visual.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### delete()
Delete the visual.

#### Syntax
```js
visualObject.delete();
```

#### Parameters
None

#### Returns
void
