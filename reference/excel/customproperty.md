# CustomProperty Object (JavaScript API for Excel)

Represents a custom property.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|key|string|Gets the key of the custom property. Read only. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Gets the value type of the custom property. Read only. Read-only. Possible values are: Number, Boolean, Date, String, Float.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|value|object|Gets or sets the value of the custom property.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the custom property.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### delete()
Deletes the custom property.

#### Syntax
```js
customPropertyObject.delete();
```

#### Parameters
None

#### Returns
void
