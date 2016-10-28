# Hyperlink Object (JavaScript API for Visio)

_Visio Online_

Represents the Hyperlink.

## Properties

| Property	   | Type	|Description| Req. Set| Feedback|
|:---------------|:--------|:----------|:----|:---|
|address|string|Gets the address for a shape's Hyperlink object, the address to which the hyperlink navigates. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlink-address)|
|description|string|Gets the description of a hyperlink. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlink-description)|
|name|string|Hyperlink name. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlink-name)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set| Feedback|
|:---------------|:--------|:----------|:----|:---|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlink-load)|

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
### Property access examples
```js
Visio.run(function (ctx) { 
	var activePage = ctx.document.getActivePage();
	var shapeName = "Sample Name";
	var shape = activePage.shapes.getItem(shapeName);
	var hyperlink = shape.hyperlinks.getItem(0);
	return ctx.sync().then(function() {
console.log(hyperlink.description);
console.log(hyperlink.address);
 });
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```