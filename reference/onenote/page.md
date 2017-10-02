# Page Object (JavaScript API for OneNote)

_Applies to: OneNote Online_   


Represents a OneNote page.

## Properties

| Property	   | Type	|Description|Feedback|
|:---------------|:--------|:----------|:-------|
|clientUrl|string|The client url of the page. Read only Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-clientUrl)|
|id|string|Gets the ID of the page. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-id)|
|pageLevel|int|Gets or sets the indentation level of the page.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-pageLevel)|
|title|string|Gets or sets the title of the page.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-title)|
|webUrl|string|The web url of the page. Read only Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-webUrl)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|contents|[PageContentCollection](pagecontentcollection.md)|The collection of PageContent objects on the page. Read only Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-contents)|
|inkAnalysisOrNull|[InkAnalysis](inkanalysis.md)|Text interpretation for the ink on the page. Returns null if there is no ink analysis information. Read only. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-inkAnalysisOrNull)|
|parentSection|[Section](section.md)|Gets the section that contains the page. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-parentSection)|

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|[addOutline(left: double, top: double, html: String)](#addoutlineleft-double-top-double-html-string)|[Outline](outline.md)|Adds an Outline to the page at the specified position.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-addOutline)|
|[copyToSection(destinationSection: Section)](#copytosectiondestinationsection-section)|[Page](page.md)|Copies this page to specified section.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-copyToSection)|
|[getRestApiId()](#getRestApiId)|string|Gets the ID that is compatible with the REST API.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-getRestApiId)|
|[insertPageAsSibling(location: string, title: string)](#insertpageassiblinglocation-string-title-string)|[Page](page.md)|Inserts a new page before or after the current page.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-insertPageAsSibling)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-load)|

## Method Details


### addOutline(left: double, top: double, html: String)
Adds an Outline to the page at the specified position.

#### Syntax
```js
pageObject.addOutline(left, top, html);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|left|double|The left position of the top, left corner of the Outline.|
|top|double|The top position of the top, left corner of the Outline.|
|html|String|An HTML string that describes the visual presentation of the Outline. See [supported HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) for the OneNote add-ins JavaScript API.|

#### Returns
[Outline](outline.md)

#### Examples
```js
OneNote.run(function (context) {

    // Gets the active page.
    var page = context.application.getActivePage();

    // Queue a command to add an outline with given html. 
    var outline = page.addOutline(200, 200,
"<p>Images and a table below:</p> \
 <img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\"> \
 <img src=\"http://imagenes.es.sftcdn.net/es/scrn/6653000/6653659/microsoft-onenote-2013-01-535x535.png\"> \
 <table> \
   <tr> \
     <td>Jill</td> \
     <td>Smith</td> \
     <td>50</td> \
   </tr> \
   <tr> \
     <td>Eve</td> \
     <td>Jackson</td> \
     <td>94</td> \
   </tr> \
 </table>"     
        );

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
		});
});
```


### copyToSection(destinationSection: Section)
Copies this page to specified section.

#### Syntax
```js
pageObject.copyToSection(destinationSection);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|destinationSection|Section|The section to copy this page to.|

#### Returns
[Page](page.md)

#### Examples
```js
OneNote.run(function(ctx) {
	var app = ctx.application;
	
	// Gets the active notebook.
	var notebook = app.getActiveNotebook();
	
	// Gets the active page.
	var page = app.getActivePage();
	
	// Queue a command to load sections under the notebook.
	notebook.load('sections');
	
	var newPage;
	
	// Run the queued commands, and return a promise to indicate task completion.
	return ctx.sync()
		.then(function() {
			var section = notebook.sections.items[0];
			
			// copy page to the section.
			newPage = page.copyToSection(section);
			newPage.load('id');
			return ctx.sync();
		})
		.then(function() {
			console.log(newPage.id);
		});
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

### getRestApiId()
Gets the ID that is compatible with the REST API.

#### Syntax
```js
pageObject.getRestApiId();
```

#### Parameters
None

#### Returns
string

#### Examples
```js

OneNote.run(function(ctx){
    // Get the current page.         
    var page = ctx.application.getActivePage();
    var restApiId = page.getRestApiId();

    return ctx.sync().
        then(function(){
            console.log("The REST API ID is " + restApiId.value);
            // Note that the REST API ID isn't all you need to interact with the OneNote REST API. For SharePoint notebooks, the notebook baseUrl should be used to talk to the OneNote REST API according to [OneNote Development Blog](https://blogs.msdn.microsoft.com/onenotedev/2015/06/11/and-sharepoint-makes-three/)
            // (this is only required for SharePoint notebooks, baseUrl will be null for OneDrive notebooks)
        });
});
```

### insertPageAsSibling(location: string, title: string)
Inserts a new page before or after the current page.

#### Syntax
```js
pageObject.insertPageAsSibling(location, title);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|location|string|The location of the new page relative to the current page.  Possible values are: Before, After|
|title|string|The title of the new page.|

#### Returns
[Page](page.md)

#### Examples
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Queue a command to add a new page after the active page. 
    var newPage = activePage.insertPageAsSibling("After", "Next Page");

    // Queue a command to load the newPage to access its data.
    context.load(newPage);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log("page is created with title: " + newPage.title);
        });
})
.catch(function(error) {
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
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
### Property access examples

**contents**
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Queue a command to add a new page after the active page. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            for(var i=0; i < pageContents.items.length; i++)
            {
                var pageContent = pageContents.items[i];
                if (pageContent.type == "Outline")
                {
                    console.log("Found an outline");
                }
                else if (pageContent.type == "Image")
                {
                    console.log("Found an image");
                }
                else if (pageContent.type == "Other")
                {
                    console.log("Found a type not supported yet.");
                }
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

**webUrl**
```js
OneNote.run(function (context) {

	var app = context.application;
	
	// Gets the active page.
	var page = app.getActivePage();
	
	// Queue a command to load the webUrl of the page.
	page.load("webUrl");
	
	// Run the queued commands, and return a promise to indicate task completion.
	return context.sync()
		.then(function() {
			console.log(page.webUrl);
		});
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

**inkAnalysisOrNull**
```js
OneNote.run(function (ctx) {		
	var app = ctx.application;
	
	// Gets the active page.
	var page = app.getActivePage();
	
	// Load ink words
	page.load('inkAnalysisOrNull/paragraphs/lines/words');
	
	return ctx.sync()
		.then(function() {
			if (!page.inkAnalysisOrNull.isNull)
				console.log(page.inkAnalysisOrNull.paragraphs.length);
		});
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

