# ShapeDataItem Object (JavaScript API for Visio)

_Visio Online_

Represents the ShapeDataItem.

## Properties

| Property	   | Type	|Description| Req. Set| Feedback|
|:---------------|:--------|:----------|:----|:---|
|label|string|Label. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItem-label)|
|name|string|Name. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItem-name)|
|value|string|Value. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItem-value)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set| Feedback|
|:---------------|:--------|:----------|:----|:---|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItem-load)|

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
	var shape = activePage.shapes.getItem(0);
        var shapeDataItems = shape.shapeDataItems;
        shapeDataItems.load('items');
        return ctx.sync().then(function() {
            for (var i = 0; i < shapeDataItems.items.length; i++)
            {
                console.log(shapeDataItems.items[i].label);
                console.log(shapeDataItems.items[i].value);
            }
        });
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
