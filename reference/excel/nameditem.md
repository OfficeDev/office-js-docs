# NamedItem object (JavaScript API for Excel)

Represents a defined name for a range of cells or value. Names can be primitive named objects (as seen in the type below), range object, reference to a range. This object can be used to obtain range object associated with names.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|name|string|The name of the object. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Indicates what type of reference is associated with the name. Read-only. Possible values are: String, Integer, Double, Boolean, Range.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|value|object|Represents the formula that the name is defined to refer to. E.g. =Sheet14!$B$2:$H$12, =4.75, etc. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|Specifies whether the object is visible or not.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getRange()](#getrange)|[Range](range.md)|Returns the range object that is associated with the name. Throws an exception if the named item's type is not a range.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getRange()
Returns the range object that is associated with the name. Throws an exception if the named item's type is not a range.

#### Syntax
```js
namedItemObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

Returns the Range object that is associated with the name. `null` if the name is not of the type `Range`. Note: This API currently supports only the Workbook scoped items.**

```js
Excel.run(function (ctx) { 
	var names = ctx.workbook.names;
	var range = names.getItem('MyRange').getRange();
	range.load('address');
	return ctx.sync().then(function() {
			console.log(range.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


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

```js
Excel.run(function (ctx) { 
	var names = ctx.workbook.names;
	var namedItem = names.getItem('MyRange');
	namedItem.load('type');
	return ctx.sync().then(function() {
			console.log(namedItem.type);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
