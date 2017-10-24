# PageContentCollection Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  

Represents the contents of a page, as a collection of PageContent objects.

To provide feedback on this API, you can [file an issue in GitHub](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection).

## Properties

| Property	 | Type	   |Description|
|:---------------|:--------|:----------|
|count|int|Returns the number of page contents in the collection. Read-only.|
|items|[PageContent[]](pagecontent.md)|A collection of pageContent objects. Read-only.|

_See [property access examples](#property-access-examples)_.

## Relationships

None


## Methods

| Method	 |Return Type |Description| 
|:---------------|:--------|:----------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[PageContent](pagecontent.md)|Gets a PageContent object by ID or by its index in the collection. Read-only.|
|[getItemAt(index: number)](#getitematindex-number)|[PageContent](pagecontent.md)|Gets a page content on its position in the collection.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.|

## Method details

### getItem(index: number or string)

Gets a PageContent object by ID or by its index in the collection. Read-only.

#### Syntax

```js
pageContentCollectionObject.getItem(index);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number or string|The ID of the PageContent object, or the index location of the PageContent object in the collection.|

#### Returns

[PageContent](pagecontent.md)

<br/>

### getItemAt(index: number)

Gets a page content on its position in the collection.

#### Syntax

```js
pageContentCollectionObject.getItemAt(index);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns

[PageContent](pagecontent.md)

#### Examples

```js
OneNote.run(function (context) {

	var page = context.application.getActivePage();
	var pageContents = page.contents;
	var firstPageContent = pageContents.getItemAt(0);
	firstPageContent.load('type');

	// Run the queued commands, and return a promise to indicate task completion.
	return context.sync()
		.then(function () {
			console.log("The first page content item is of type: " + firstPageContent.type);
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

**items**

```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Queue a command to load the type of each pageContent.
    pageContents.load("type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            $.each(pageContents.items, function(index, pageContent) {
                console.log("PageContent type: " + pageContent.type);
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

**traverse for outlines**

```js
OneNote.run(function (context) {
   var page = context.application.getActivePage();
   var pageContents = page.contents;
   pageContents.load('type');
   var outlines = [];
   return context.sync()
	   .then(function () {	  
			  $.each(pageContents.items, function (index, pageContent) {
					 console.log(pageContent.type);
					 if (pageContent.type === 'Outline') {
						   outlines.push(pageContent);
					 }
			  });
			  $.each(outlines, function (index, outline) {
					 outline.load("id,paragraphs,paragraphs/type");
			  });
			  return context.sync();
	   })
	   .then(function () {
			  $.each(outlines, function (index, outline) {
					 console.log("An outline was found with id : " + outline.id);
			  });
			  return Promise.resolve(outlines);
	   });
});
```

<br/>
