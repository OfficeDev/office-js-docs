# PageCollection Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  

Represents a collection of pages.

To provide feedback on this API, you can [file an issue in GitHub](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection).

## Properties

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|count|int|Returns the number of pages in the collection. Read-only.|
|items|[Page[]](page.md)|A collection of page objects. Read-only.|

_See [property access examples](#property-access-examples)_.

## Relationships

None


## Methods

| Method		   | Return Type	|Description| 
|:---------------|:--------|:----------|
|[getByTitle(title: string)](#getbytitletitle-string)|[PageCollection](pagecollection.md)|Gets the collection of pages with the specified title.|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[Page](page.md)|Gets a page by ID or by its index in the collection. Read-only.|
|[getItemAt(index: number)](#getitematindex-number)|[Page](page.md)|Gets a page on its position in the collection.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.|

## Method details

### getByTitle(title: string)

Gets the collection of pages with the specified title.

#### Syntax

```js
pageCollectionObject.getByTitle(title);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|title|string|The title of the page.|

#### Returns

[PageCollection](pagecollection.md)

#### Examples

```js
OneNote.run(function (context) {

    // Get all the pages in the current section.
    var allPages = context.application.getActiveSection().pages;

    // Queue a command to load the pages. 
    // For best performance, request specific properties.
    allPages.load("id"); 

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Get the sections with the specified name.
            var todoPages = allPages.getByTitle("Todo list");

            // Queue a command to load the section. 
            // For best performance, request specific properties.
            todoPages.load("id,title"); 

            return context.sync()
                .then(function () {

                    // Iterate through the collection or access items individually by index.
                    if (todoPages.items.length > 0) {
                        console.log("Page title: " + todoPages.items[0].title);
                        console.log("Page ID: " + todoPages.items[0].id);
                    }
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

### getItem(index: number or string)

Gets a page by ID or by its index in the collection. Read-only.

#### Syntax

```js
pageCollectionObject.getItem(index);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number or string|The ID of the page, or the index location of the page in the collection.|

#### Returns

[Page](page.md)

<br/>

### getItemAt(index: number)

Gets a page on its position in the collection.

#### Syntax

```js
pageCollectionObject.getItemAt(index);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns

[Page](page.md)

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
    
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
    
    // Queue a command to load the id and title for each page.            
    pages.load('id,title');
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Display the properties.
            $.each(pages.items, function(index, page) {
                console.log(page.title);
                console.log(page.id);
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

