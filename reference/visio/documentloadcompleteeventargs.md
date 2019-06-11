# DocumentLoadCompleteEventArgs Object (JavaScript API for Visio)

Applies to: _Visio Online_

Provides information about the success or failure of the DocumentLoadComplete event.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|success|bool|Gets the success or failure of the DocumentLoadComplete event.|

_See property access [examples.](#property-access-examples)_

## Relationships
None

## Methods
None

### Property access examples
```js
Visio.run(session, function (ctx) { 
  var document1 = ctx.document;
	     	eventResult1 = document1.onDocumentLoadComplete.add(
			  function (args){
			       console.log("Document Loaded");
			});

	return ctx.sync().then(function () {
		   console.log("Success");
		});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
