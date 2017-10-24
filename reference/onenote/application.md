# Application Object (JavaScript API for OneNote)

_Applies to: OneNote Online_

Represents the top-level object that contains all globally addressable OneNote objects such as notebooks, the active notebook, and the active section.

To provide feedback on this API, you can [file an issue in GitHub](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application).

## Properties

None

## Relationships

| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|notebooks|[NotebookCollection](notebookcollection.md)|Gets the collection of notebooks that are open in the OneNote application instance. In OneNote Online, only one notebook at a time is open in the application instance. Read-only.|

## Methods

| Method		   | Return Type	|Description| 
|:---------------|:--------|:----------|
|[getActiveNotebook()](#getactivenotebook)|[Notebook](notebook.md)|Gets the active notebook if one exists. If no notebook is active, throws ItemNotFound.|
|[getActiveNotebookOrNull()](#getactivenotebookornull)|[Notebook](notebook.md)|Gets the active notebook if one exists. If no notebook is active, returns null.|
|[getActiveOutline()](#getactiveoutline)|[Outline](outline.md)|Gets the active outline if one exists. If no outline is active, throws ItemNotFound.|
|[getActiveOutlineOrNull()](#getactiveoutlineornull)|[Outline](outline.md)|Gets the active outline if one exists; otherwise, returns null.|
|[getActivePage()](#getactivepage)|[Page](page.md)|Gets the active page if one exists. If no page is active, throws ItemNotFound.|
|[getActivePageOrNull()](#getactivepageornull)|[Page](page.md)|Gets the active page if one exists. If no page is active, returns null.|
|[getActiveSection()](#getactivesection)|[Section](section.md)|Gets the active section if one exists. If no section is active, throws ItemNotFound.|
|[getActiveSectionOrNull()](#getactivesectionornull)|[Section](section.md)|Gets the active section if one exists. If no section is active, returns null.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.|
|[navigateToPage(page: Page)](#navigatetopagepage-page)|void|Opens the specified page in the application instance.|
|[navigateToPageWithClientUrl(url: string)](#navigatetopagewithclienturlurl-string)|[Page](page.md)|Gets the specified page, and opens it in the application instance.|

## Method details

### getActiveNotebook()
Gets the active notebook if one exists. If no notebook is active, throws ItemNotFound.

#### Syntax

```js
applicationObject.getActiveNotebook();
```

#### Parameters

None

#### Returns

[Notebook](notebook.md)

#### Examples

```js
OneNote.run(function (context) {
        
    // Get the active notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Notebook name: " + notebook.name);
            console.log("Notebook ID: " + notebook.id);
            
        });
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>

### getActiveNotebookOrNull()

Gets the active notebook if one exists. If no notebook is active, returns null.

#### Syntax

```js
applicationObject.getActiveNotebookOrNull();
```

#### Parameters
None

#### Returns

[Notebook](notebook.md)

#### Examples

```js
OneNote.run(function (context) {

    // Get the active notebook.
    var notebook = context.application.getActiveNotebookOrNull();

    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id,name');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // check if active notebook is set.
            if (!notebook.isNull) {
                console.log("Notebook name: " + notebook.name);
                console.log("Notebook ID: " + notebook.id);
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

<br/>

### getActiveOutline()

Gets the active outline if one exists. If no outline is active, throws ItemNotFound.

#### Syntax

```js
applicationObject.getActiveOutline();
```

#### Parameters

None

#### Returns

[Outline](outline.md)

#### Examples

```js
OneNote.run(function (context) {

    // get active outline.
    var outline = context.application.getActiveOutline();

    // Queue a command to load the id of the outline.         
    outline.load('id');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Show some properties.
            console.log("outline id: " + outline.id);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

<br/>

### getActiveOutlineOrNull()

Gets the active outline if one exists; otherwise, returns null.

#### Syntax

```js
applicationObject.getActiveOutlineOrNull();
```

#### Parameters

None

#### Returns

[Outline](outline.md)

#### Examples

```js
OneNote.run(function (context) {

    // get active outline.
    var outline = context.application.getActiveOutlineOrNull();

    // Queue a command to load the id of the outline.         
    outline.load('id');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            if (!outline.isNull) {
                console.log("outline id: " + outline.id);
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

<br/>

### getActivePage()

Gets the active page if one exists. If no page is active, throws ItemNotFound.

#### Syntax

```js
applicationObject.getActivePage();
```

#### Parameters

None

#### Returns

[Page](page.md)

#### Examples

```js
OneNote.run(function (context) {
        
    // Get the active page.
    var page = context.application.getActivePage();
            
    // Queue a command to load the page. 
    // For best performance, request specific properties.           
    page.load('id,title');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Page title: " + page.title);
            console.log("Page ID: " + page.id);
            
        });
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>

### getActivePageOrNull()

Gets the active page if one exists. If no page is active, returns null.

#### Syntax

```js
applicationObject.getActivePageOrNull();
```

#### Parameters

None

#### Returns

[Page](page.md)

#### Examples

```js
OneNote.run(function (context) {

    // Get the active page.
    var page = context.application.getActivePageOrNull();

    // Queue a command to load the page. 
    // For best performance, request specific properties.           
    page.load('id,title');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            if (!page.isNull) {
                // Show some properties.
                console.log("Page title: " + page.title);
                console.log("Page ID: " + page.id);
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

<br/>

### getActiveSection()

Gets the active section if one exists. If no section is active, throws ItemNotFound.

#### Syntax

```js
applicationObject.getActiveSection();
```

#### Parameters

None

#### Returns

[Section](section.md)

#### Examples

```js
OneNote.run(function (context) {
        
    // Get the active section.
    var section = context.application.getActiveSection();
            
    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Section name: " + section.name);
            console.log("Section ID: " + section.id);
            
        });
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>

### getActiveSectionOrNull()

Gets the active section if one exists. If no section is active, returns null.

#### Syntax

```js
applicationObject.getActiveSectionOrNull();
```

#### Parameters

None

#### Returns

[Section](section.md)

#### Examples

```js
OneNote.run(function (context) {

    // Get the active section.
    var section = context.application.getActiveSectionOrNull();

    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load('id,name');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if (!section.isNull) {
                // Show some properties.
                console.log("Section name: " + section.name);
                console.log("Section ID: " + section.id);
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

<br/>

### load(param: object)

Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.

#### Syntax

```js
object.load(param);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns

Void

<br/>

### navigateToPage(page: Page)

Opens the specified page in the application instance.

#### Syntax

```js
applicationObject.navigateToPage(page);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|page|Page|The page to open.|

#### Returns

Void

#### Examples

```js        
OneNote.run(function (context) {
        
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
            
    // Queue a command to load the pages. 
    // For best performance, request specific properties.           
    pages.load('id');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // This example loads the first page in the section.
            var page = pages.items[0];
                        
            // Open the page in the application.                    
            context.application.navigateToPage(page);
                    
            // Run the queued command.
            return context.sync();
        });
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>

### navigateToPageWithClientUrl(url: string)

Gets the specified page, and opens it in the application instance.

#### Syntax

```js
applicationObject.navigateToPageWithClientUrl(url);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|url|string|The client URL of the page to open.|

#### Returns

[Page](page.md)

#### Examples

```js
OneNote.run(function (context) {

    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;

    // Queue a command to load the pages. 
    // For best performance, request specific properties.           
    pages.load('clientUrl');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // This example loads the first page in the section.
            var page = pages.items[0];

            // Open the page in the application.                    
            context.application.navigateToPageWithClientUrl(page.clientUrl);

            // Run the queued command.
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

<br/>
