# Paragraph Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  

A container for the visible content on a page. A Paragraph can contain any one ParagraphType type of content.

To provide feedback on this API, you can [file an issue in GitHub](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph).

## Properties

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|Gets the ID of the Paragraph object. Read-only.|
|type|string|Gets the type of the Paragraph object. Read-only. Possible values are RichText, Image, Table, Other.|

_See [property access examples](#property-access-examples)_.

## Relationships

| Relationship | Type	|Description| 
|:---------------|:--------|:----------|
|image|[Image](image.md)|Gets the Image object in the Paragraph. Throws an exception if ParagraphType is not Image. Read-only.|
|inkWords|[InkWordCollection](inkwordcollection.md)|Gets the Ink collection in the Paragraph. Throws an exception if ParagraphType is not Ink. Read-only.|
|outline|[Outline](outline.md)|Gets the Outline object that contains the Paragraph. Read-only.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|The collection of paragraphs under this paragraph. Read-only.|
|parentParagraph|[Paragraph](paragraph.md)|Gets the parent paragraph object. Throws an exception if a parent paragraph does not exist. Read-only.|
|parentParagraphOrNull|[Paragraph](paragraph.md)|Gets the parent paragraph object. Returns null if a parent paragraph does not exist. Read-only.|
|parentTableCell|[TableCell](tablecell.md)|Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, throws ItemNotFound. Read-only.|
|parentTableCellOrNull|[TableCell](tablecell.md)|Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, returns null. Read-only.|
|richText|[RichText](richtext.md)|Gets the RichText object in the Paragraph. Throws an exception if ParagraphType is not RichText. Read-only.|
|table|[Table](table.md)|Gets the Table object in the Paragraph. Throws an exception if ParagraphType is not Table. Read-only.|

## Methods

| Method		   | Return Type	|Description| 
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Deletes the paragraph.|
|[insertHtmlAsSibling(insertLocation: string, html: string)](#inserthtmlassiblinginsertlocation-string-html-string)|void|Inserts the specified HTML content.|
|[insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: double, height: double)](#insertimageassiblinginsertlocation-string-base64encodedimage-string-width-double-height-double)|[Image](image.md)|Inserts the image at the specified insert location.|
|[insertRichTextAsSibling(insertLocation: string, paragraphText: string)](#insertrichtextassiblinginsertlocation-string-paragraphtext-string)|[RichText](richtext.md)|Inserts the paragraph text at the specified insert location.|
|[insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values: string[][])](#inserttableassiblinginsertlocation-string-rowcount-number-columncount-number-values-string)|[Table](table.md)|Adds a table with the specified number of rows and columns before or after the current paragraph.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.|

## Method details

### delete()

Deletes the paragraph.

#### Syntax

```js
paragraphObject.delete();
```

#### Parameters

None

#### Returns

Void

#### Examples

```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
	
    var paragraphs = pageContent.outline.paragraphs;
	
	var firstParagraph = paragraphs.getItemAt(0);
	
    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
			
            // Queue a command to delete the first paragraph                 
            firstParagraph.delete();
			
			// Run the command to delete it
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

### insertHtmlAsSibling(insertLocation: string, html: string)

Inserts the specified HTML content.

#### Syntax

```js
paragraphObject.insertHtmlAsSibling(insertLocation, html);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|insertLocation|string|The location of new contents relative to the current Paragraph. Possible values are Before, After.|
|html|string|An HTML string that describes the visual presentation of the content. See [supported HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) for the OneNote add-ins JavaScript API.|

#### Returns

Void

#### Examples

```js
OneNote.run(function (context) {

	// Get the collection of pageContent items from the page.
	var pageContents = context.application.getActivePage().contents;

	// Get the first PageContent on the page
	// Assuming its an outline, get the outline's paragraphs.
	var pageContent = pageContents.getItemAt(0);
	var paragraphs = pageContent.outline.paragraphs;
	var firstParagraph = paragraphs.getItemAt(0);

	// Queue a command to load the id and type of the first paragraph
	firstParagraph.load("id,type");

	// Run the queued commands, and return a promise to indicate task completion.
	return context.sync()
		.then(function () {

			// Queue commands to insert before and after the first paragraph
			firstParagraph.insertHtmlAsSibling("Before", "<p>ContentBeforeFirstParagraph</p>");
			firstParagraph.insertHtmlAsSibling("After", "<p>ContentAfterFirstParagraph</p>");
			
			// Run the command to run inserts
			return context.sync();
		});
))
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>

### insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: double, height: double)

Inserts the image at the specified insert location.

#### Syntax

```js
paragraphObject.insertImageAsSibling(insertLocation, base64EncodedImage, width, height);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|insertLocation|string|The location of the table relative to the current Paragraph. Possible values are Before, After.|
|base64EncodedImage|string|HTML string to append.|
|width|double|Optional. Width in the unit of Points. The default value is null, and image width will be respected.|
|height|double|Optional. Height in the unit of Points. The default value is null, and image height will be respected.|

#### Returns

[Image](image.md)

#### Examples

```js
OneNote.run(function (context) {

	// Get the collection of pageContent items from the page.
	var pageContents = context.application.getActivePage().contents;

	// Get the first PageContent on the page
	// Assuming its an outline, get the outline's paragraphs.
	var pageContent = pageContents.getItemAt(0);
	var paragraphs = pageContent.outline.paragraphs;
	var firstParagraph = paragraphs.getItemAt(0);

	// Queue a command to load the id and type of the first paragraph
	firstParagraph.load("id,type");

	// Run the queued commands, and return a promise to indicate task completion.
	return context.sync()
		.then(function () {

			// Queue commands to insert before and after the first paragraph
			firstParagraph.insertImageAsSibling("Before", "R0lGODlhDwAPAKECAAAAzMzM/////wAAACwAAAAADwAPAAACIISPeQHsrZ5ModrLlN48CXF8m2iQ3YmmKqVlRtW4MLwWACH+H09wdGltaXplZCBieSBVbGVhZCBTbWFydFNhdmVyIQAAOw==");
			firstParagraph.insertImageAsSibling("After", "R0lGODlhDwAPAKECAAAAzMzM/////wAAACwAAAAADwAPAAACIISPeQHsrZ5ModrLlN48CXF8m2iQ3YmmKqVlRtW4MLwWACH+H09wdGltaXplZCBieSBVbGVhZCBTbWFydFNhdmVyIQAAOw==");
			
			// Run the command to insert images
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

### insertRichTextAsSibling(insertLocation: string, paragraphText: string)

Inserts the paragraph text at the specified insert location.

#### Syntax

```js
paragraphObject.insertRichTextAsSibling(insertLocation, paragraphText);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|insertLocation|string|The location of the table relative to the current Paragraph. Possible values are Before, After.|
|paragraphText|string|HTML string to append.|

#### Returns

[RichText](richtext.md)

#### Examples

```js
OneNote.run(function (context) {

	// Get the collection of pageContent items from the page.
	var pageContents = context.application.getActivePage().contents;

	// Get the first PageContent on the page
	// Assuming its an outline, get the outline's paragraphs.
	var pageContent = pageContents.getItemAt(0);
	var paragraphs = pageContent.outline.paragraphs;
	var firstParagraph = paragraphs.getItemAt(0);

	// Queue a command to load the id and type of the first paragraph
	firstParagraph.load("id,type");

	// Run the queued commands, and return a promise to indicate task completion.
	return context.sync()
		.then(function () {

			// Queue commands to insert before and after the first paragraph
			firstParagraph.insertRichTextAsSibling("Before", "Text Appears Before Paragraph");
			firstParagraph.insertRichTextAsSibling("After", "Text Appears After Paragraph");
			
			// Run the command to insert text contents
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

### insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values: string[][])

Adds a table with the specified number of rows and columns before or after the current paragraph.

#### Syntax

```js
paragraphObject.insertTableAsSibling(insertLocation, rowCount, columnCount, values);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|insertLocation|string|The location of the table relative to the current Paragraph. Possible values are Before, After.|
|rowCount|number|The number of rows in the table.|
|columnCount|number|The number of columns in the table.|
|values|string[][]|Optional. Optional 2D array. Cells are filled if the corresponding strings are specified in the array.|

#### Returns

[Table](table.md)

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
|param|object|Optional. Accepts parameter and relationship names as a delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns

Void

<br/>

### Property access examples

**id and type**

```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;
    
    // Queue a command to load the outline property of each pageContent.
    pageContents.load("outline");
        
    // Get the first PageContent on the page, and then get its Outline.
    var pageContent = pageContents._GetItem(0);
    var paragraphs = pageContent.outline.paragraphs;
            
    // Queue a command to load the id and type of each paragraph.
    paragraphs.load("id,type");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Write the text.                  
            $.each(paragraphs.items, function(index, paragraph) {
                console.log("Paragraph type: " + paragraph.type);
                console.log("Paragraph ID: " + paragraph.id);
            });
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

**paragraphs**

```js
OneNote.run(function(context) {
	var app = context.application;
	
	// Gets the active outline
	var outline = app.getActiveOutline();
	
	// load nested paragraphs and their types.
	outline.load("paragraphs/type");
	
	return context.sync().then(function () {
		var paragraphs = outline.paragraphs.items;
		
		var promise;
		// for each nested paragraphs, load tables only
		for (var i = 0; i < paragraphs.length; i++) {
			var paragraph = paragraphs[i];
			if (paragraph.type == "Table") {
				paragraph.load("table/id");
				promise =  context.sync().then(function() {
					console.log(paragraph.table.id);
				});
			}
		}
		return promise;
	})
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>
