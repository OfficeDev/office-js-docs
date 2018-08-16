# Application object (JavaScript API for Visio)

Applies to: _Visio Online_

Represents the Application.

## Properties

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|showToolbars|bool|Show or Hide the standard toolbars.|
|showBorders|bool|Show or Hide the iframe applicaiton borders.|

_See property access [examples.](#property-access-examples)_

## Relationships
None

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[showToolbar(ToolBarType: ToolBarType, Visibility: bool)](#showtoolbartoolbartype-toolbartype-visibility-bool)|void| Sets the visibility of a specific toolbar in the application.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details

### showToolbar(ToolBarType: ToolBarType, Visibility: bool)
Sets the visibility of a specific toolbar in the application.

#### Syntax
```js
applicationObject.showToolbar(ToolBarType, Visibility);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|ToolBarType|ToolBarType| 1 - Command Bar, 2 - Page Navigation Bar, 3 - Status Bar|
|Visibility|bool|Visibility of the toolbar – True, False|

#### Returns
void

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
Visio.run(session, function (ctx) { 
	var application = ctx.document.application;
	application.showToolbars = false;
	application.showBorders = false;
	return ctx.sync();
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
