# ContentControlCollection Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Contains a collection of [contentControl](contentControl.md) objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[ContentControl[]](contentcontrol.md)|A collection of contentControl objects. Read-only.|[1.1](../requirement-sets/word-api-requirement-sets.md)|

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getById(id: number)](#getbyidid-number)|[ContentControl](contentcontrol.md)|Gets a content control by its identifier. Throws if there isn't a content control with the identifier in this collection.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|[getByIdOrNullObject(id: number)](#getbyidornullobjectid-number)|[ContentControl](contentcontrol.md)|Gets a content control by its identifier. Returns a null object if there isn't a content control with the identifier in this collection.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getByTag(tag: string)](#getbytagtag-string)|[ContentControlCollection](contentcontrolcollection.md)|Gets the content controls that have the specified tag.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|[getByTitle(title: string)](#getbytitletitle-string)|[ContentControlCollection](contentcontrolcollection.md)|Gets the content controls that have the specified title.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|[getByTypes(types: ContentControlType[])](#getbytypestypes-contentcontroltype)|[ContentControlCollection](contentcontrolcollection.md)|Gets the content controls that have the specified types andor subtypes.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getFirst()](#getfirst)|[ContentControl](contentcontrol.md)|Gets the first content control in this collection. Throws if this collection is empty.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getFirstOrNullObject()](#getfirstornullobject)|[ContentControl](contentcontrol.md)|Gets the first content control in this collection. Returns a null object if this collection is empty.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getItem(index: number)](#getitemindex-number)|[ContentControl](contentcontrol.md)|Gets a content control by its index in the collection.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/word-api-requirement-sets.md)|

## Method Details


### getById(id: number)
Gets a content control by its identifier. Throws if there isn't a content control with the identifier in this collection.

#### Syntax
```js
contentControlCollectionObject.getById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|number|Required. A content control identifier.|

#### Returns
[ContentControl](contentcontrol.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

	// Create a proxy object for the content control that contains a specific id.
	var contentControl = context.document.contentControls.getById(30086310);

	// Queue a command to load the text property for a content control.
	context.load(contentControl, 'text');

	// Synchronize the document state by executing the queued commands,
	// and return a promise to indicate task completion.
	return context.sync().then(function () {
		console.log('The content control with that Id has been found in this document.');
	});
})
.catch(function (error) {
	console.log('Error: ' + JSON.stringify(error));
	if (error instanceof OfficeExtension.Error) {
		console.log('Debug info: ' + JSON.stringify(error.debugInfo));
	}
});
```


#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

	// Create a proxy object for the content control that contains a specific id.
    var contentControl = context.document.contentControls.getByIdOrNullObject(30086310);

    // Queue a command to load the text property for a content control.
    context.load(contentControl, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControl.isNullObject) {
            console.log('There is no content control with that ID.')
        } else {
            console.log('The content control with that ID has been found in this document.');
        }
    });
})
.catch(function (error) {
	console.log('Error: ' + JSON.stringify(error));
	if (error instanceof OfficeExtension.Error) {
		console.log('Debug info: ' + JSON.stringify(error.debugInfo));
	}
});
```



### getByIdOrNullObject(id: number)
Gets a content control by its identifier. Returns a null object if there isn't a content control with the identifier in this collection.

#### Syntax
```js
contentControlCollectionObject.getByIdOrNullObject(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|number|Required. A content control identifier.|

#### Returns
[ContentControl](contentcontrol.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

	// Create a proxy object for the content control that contains a specific id.
    var contentControl = context.document.contentControls.getByIdOrNullObject(30086310);

    // Queue a command to load the text property for a content control.
    context.load(contentControl, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControl.isNullObject) {
            console.log('There is no content control with that ID.')
        } else {
            console.log('The content control with that ID has been found in this document.');
        }
    });
})
.catch(function (error) {
	console.log('Error: ' + JSON.stringify(error));
	if (error instanceof OfficeExtension.Error) {
		console.log('Debug info: ' + JSON.stringify(error.debugInfo));
	}
});
```



### getByTag(tag: string)
Gets the content controls that have the specified tag.

#### Syntax
```js
contentControlCollectionObject.getByTag(tag);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|tag|string|Required. A tag set on a content control.|

#### Returns
[ContentControlCollection](contentcontrolcollection.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content controls collection that contains a specific tag.
    var contentControlsWithTag = context.document.contentControls.getByTag('Customer-Address');

    // Queue a command to load the text property for all of content controls with a specific tag.
    context.load(contentControlsWithTag, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTag.items.length === 0) {
            console.log("There isn't a content control with a tag of Customer-Address in this document.");
        } else {
            console.log('The first content control with the tag of Customer-Address has this text: ' + contentControlsWithTag.items[0].text);
        }

    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

*Additional information*
The [Word-Add-in-DocumentAssembly][contentControls.getByTag] sample has another example of using the getByTag method.



### getByTitle(title: string)
Gets the content controls that have the specified title.

#### Syntax
```js
contentControlCollectionObject.getByTitle(title);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|title|string|Required. The title of a content control.|

#### Returns
[ContentControlCollection](contentcontrolcollection.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content controls collection that contains a specific title.
    var contentControlsWithTitle = context.document.contentControls.getByTitle('Enter Customer Address Here');

    // Queue a command to load the text property for all of content controls with a specific title.
    context.load(contentControlsWithTitle, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTitle.items.length === 0) {
            console.log("There isn't a content control with a title of 'Enter Customer Address Here' in this document.");
        } else {
            console.log("The first content control with the title of 'Enter Customer Address Here' has this text: " + contentControlsWithTitle.items[0].text);
        }

    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

*Additional information*

The [Word-Add-in-DocumentAssembly][contentControls.getByTitle] sample has another example of using the getByTitle method.


### getByTypes(types: ContentControlType[])
Gets the content controls that have the specified types andor subtypes.

#### Syntax
```js
contentControlCollectionObject.getByTypes(types);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|types|ContentControlType[]|Required. An array of content control types and/or subtypes.|

#### Returns
[ContentControlCollection](contentcontrolcollection.md)

### getFirst()
Gets the first content control in this collection. Throws if this collection is empty.

#### Syntax
```js
contentControlCollectionObject.getFirst();
```

#### Parameters
None

#### Returns
[ContentControl](contentcontrol.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the first content control in the document.
	var contentControl = context.document.contentControls.getFirstOrNullObject();

    // Queue a command to load the text property for a content control.
    context.load(contentControl, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControl.isNullObject) {
            console.log('There are no content controls in this document.')
        } else {
            console.log('The first content control has been found in this document.');
        }
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```



### getFirstOrNullObject()
Gets the first content control in this collection. Returns a null object if this collection is empty.

#### Syntax
```js
contentControlCollectionObject.getFirstOrNullObject();
```

#### Parameters
None

#### Returns
[ContentControl](contentcontrol.md)

### getItem(index: number)
Gets a content control by its index in the collection.

#### Syntax
```js
contentControlCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|The index.|

#### Returns
[ContentControl](contentcontrol.md)

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;

    // Queue a command to load the id property for all of the content controls.
    context.load(contentControls, 'id');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to load the properties on the first content control.
            contentControls.items[0].load(  'appearance,' +
                                            'cannotDelete,' +
                                            'cannotEdit,' +
                                            'id,' +
                                            'placeHolderText,' +
                                            'removeWhenEdited,' +
                                            'title,' +
                                            'text,' +
                                            'type,' +
                                            'style,' +
                                            'tag,' +
                                            'font/size,' +
                                            'font/name,' +
                                            'font/color');

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Property values of the first content control:' +
                        '   ----- appearance: ' + contentControls.items[0].appearance +
                        '   ----- cannotDelete: ' + contentControls.items[0].cannotDelete +
                        '   ----- cannotEdit: ' + contentControls.items[0].cannotEdit +
                        '   ----- color: ' + contentControls.items[0].color +
                        '   ----- id: ' + contentControls.items[0].id +
                        '   ----- placeHolderText: ' + contentControls.items[0].placeholderText +
                        '   ----- removeWhenEdited: ' + contentControls.items[0].removeWhenEdited +
                        '   ----- title: ' + contentControls.items[0].title +
                        '   ----- text: ' + contentControls.items[0].text +
                        '   ----- type: ' + contentControls.items[0].type +
                        '   ----- style: ' + contentControls.items[0].style +
                        '   ----- tag: ' + contentControls.items[0].tag +
                        '   ----- font size: ' + contentControls.items[0].font.size +
                        '   ----- font name: ' + contentControls.items[0].font.name +
                        '   ----- font color: ' + contentControls.items[0].font.color);
            });
        }
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

The [Silly stories](https://aka.ms/sillystorywordaddin) add-in sample shows how the **load** method is used to load the content control collection with the **tag** and **title** properties.
