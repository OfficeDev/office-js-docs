# SelectionChangedEventArgs Object (JavaScript API for Visio)

Applies to: _Visio Online_

Provides information about the shape collection that raised the SelectionChanged event.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|shapeNames|string[]|Gets the array of shape names that raised the SelectionChanged event.|
|pageName|string|Gets the name of the page which has the ShapeCollection object that raised the SelectionChanged event.|

_See property access [examples.](#property-access-examples)_

## Relationships
None

## Methods
None

### Property access examples
```js
var eventResult; //Global Variable to store the EventHandlerResult returned on attaching handler.

function AttachHandler()
{
    Visio.run(session, function(ctx)
    {
        var doc = ctx.document;
        eventResult = doc.onSelectionChanged.add(
	function (args){
	       console.log("Selected Shape Name: "+args.shapeNames[0]);
	});
        return ctx.sync().then(function(){
                console.log("Handler attached");
            });
    }).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
    });

    function onSelectionChanged(args)
    {
        console.log(Date.now() + "Selection Changes Event" + JSON.stringify(args));
    }
}

function RemoveHandler()
{
    if(!eventResult || !eventResult.context)
    {
        console.log("Handler has not been attached");
        return;
    }

    Visio.run(eventResult.context, function(ctx)
    {
        eventResult.remove();
        return ctx.sync().then(function (){
            eventResult = null;
            console.log("Handler removed");
        });
    }).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
    });
}
```
