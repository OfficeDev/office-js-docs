# Document Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|saved|bool|Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.|[1.1](../requirement-sets/word-api-requirement-sets.md)|

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|body|[Body](body.md)|Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Gets the collection of content control objects in the current document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|properties|[DocumentProperties](documentproperties.md)|Gets the properties of the current document. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|sections|[SectionCollection](sectioncollection.md)|Gets the collection of section objects in the document. Read-only.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|settings|[SettingCollection](settingcollection.md)|Gets the add-in's settings in the current document. Read-only.|[Beta](../requirement-sets/word-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[deleteBookmark(name: string)](#deletebookmarkname-string)|void|Deletes a bookmark, if exists, from the document.|[Beta](../requirement-sets/word-api-requirement-sets.md)|
|[getBookmarkRange(name: string)](#getbookmarkrangename-string)|[Range](range.md)|Gets a bookmark's range. Throws if the bookmark does not exist.|[Beta](../requirement-sets/word-api-requirement-sets.md)|
|[getBookmarkRangeOrNullObject(name: string)](#getbookmarkrangeornullobjectname-string)|[Range](range.md)|Gets a bookmark's range. Returns a null object if the bookmark does not exist.|[Beta](../requirement-sets/word-api-requirement-sets.md)|
|[getSelection()](#getselection)|[Range](range.md)|Gets the current selection of the document. Multiple selections are not supported.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|[open()](#open)|void|Open the document.|[Beta](../requirement-sets/word-api-requirement-sets.md)|
|[save()](#save)|void|Saves the document. This will use the Word default file naming convention if the document has not been saved before.|[1.1](../requirement-sets/word-api-requirement-sets.md)|

## Method Details


### deleteBookmark(name: string)
Deletes a bookmark, if exists, from the document.

#### Syntax
```js
documentObject.deleteBookmark(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Required. The bookmark name, which is case-insensitive.|

#### Returns
void

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Assume the document has a header "Chapter 1" which has been made into 
    // a bookmark named "ch1".
    
    // Queue a command to delete the bookmark.
    context.document.deleteBookmark("ch1");
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync();
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


### getBookmarkRange(name: string)
Gets a bookmark's range. Throws if the bookmark does not exist.

#### Syntax
```js
documentObject.getBookmarkRange(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Required. The bookmark name, which is case-insensitive.|

#### Returns
[Range](range.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Assume the document has a header "Chapter 1" which has been made into 
    // a bookmark named "ch1".
    
    // Create a range proxy object for the bookmark.
    var bookmarkRange = context.document.getBookmarkRange("ch1");
    
    // Queue a command to load the range text property.
    context.load(bookmarkRange, 'text');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
    
        // Log "Chapter 1".
        console.log(JSON.stringify(bookmarkRange.text));
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
    
    // Create a range proxy object for the bookmark.
    var bookmarkRange = context.document.getBookmarkRangeOrNullObject("ch2");
    
    // Queue a command to load the range text property.
    context.load(bookmarkRange, 'text');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if (bookmarkRange.isNullObject) {
                console.log("There is no such bookmark.");
            }
            else {
                console.log(JSON.stringify(bookmarkRange.text));
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


### getBookmarkRangeOrNullObject(name: string)
Gets a bookmark's range. Returns a null object if the bookmark does not exist.

#### Syntax
```js
documentObject.getBookmarkRangeOrNullObject(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Required. The bookmark name, which is case-insensitive.|

#### Returns
[Range](range.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a range proxy object for the bookmark.
    var bookmarkRange = context.document.getBookmarkRangeOrNullObject("ch2");
    
    // Queue a command to load the range text property.
    context.load(bookmarkRange, 'text');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if (bookmarkRange.isNullObject) {
                console.log("There is no such bookmark.");
            }
            else {
                console.log(JSON.stringify(bookmarkRange.text));
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


### getSelection()
Gets the current selection of the document. Multiple selections are not supported.

#### Syntax
```js
documentObject.getSelection();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    var textSample = 'This is an example of the insert text method. This is a method ' + 
        'which allows users to insert text into a selection. It can insert text into a ' +
        'relative location or it can overwrite the current selection. Since the ' +
        'getSelection method returns a range object, look up the range object documentation ' +
        'for everything you can do with a selection.';
    
    // Create a range proxy object for the current selection.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert text at the end of the selection.
    range.insertText(textSample, Word.InsertLocation.end);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Inserted the text at the end of the selection.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
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
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document.
    var thisDocument = context.document;
    
    // Queue a command to load content control properties.
    context.load(thisDocument, 'contentControls/id, contentControls/text, contentControls/tag');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (thisDocument.contentControls.items.length !== 0) {
            for (var i = 0; i < thisDocument.contentControls.items.length; i++) {
                console.log(thisDocument.contentControls.items[i].id);
                console.log(thisDocument.contentControls.items[i].text);
                console.log(thisDocument.contentControls.items[i].tag);
            }
        } else {
            console.log('No content controls in this document.');
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


### open()
Open the document.

#### Syntax
```js
documentObject.open();
```

#### Parameters
None

#### Returns
void

### save()
Saves the document. This will use the Word default file naming convention if the document has not been saved before.

#### Syntax
```js
documentObject.save();
```

#### Parameters
None

#### Returns
void

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document.
    var thisDocument = context.document;

    // Queue a commmand to load the document save state (on the saved property).
    context.load(thisDocument, 'saved');    
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (thisDocument.saved === false) {
            // Queue a command to save this document.
            thisDocument.save();
            
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Saved the document');
            });
        } else {
            console.log('The document has not changed since the last save.');
        }
    });  
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
