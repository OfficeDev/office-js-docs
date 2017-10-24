# PageContent Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  

Represents a region on a page that contains top-level content types such as Outline or Image. A PageContent object can be assigned an XY position.

To provide feedback on this API, you can [file an issue in GitHub](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent).

## Properties

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|Gets the ID of the PageContent object. Read-only.|
|left|double|Gets or sets the left (X-axis) position of the PageContent object.|
|top|double|Gets or sets the top (Y-axis) position of the PageContent object.|
|type|string|Gets the type of the PageContent object. Read-only. Possible values are Outline, Image, Other.|

## Relationships

| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|image|[Image](image.md)|Gets the Image in the PageContent object. Throws an exception if PageContentType is not Image. Read-only.|
|ink|[FloatingInk](floatingink.md)|Gets the ink in the PageContent object. Throws an exception if PageContentType is not Ink. Read-only.|
|outline|[Outline](outline.md)|Gets the Outline in the PageContent object. Throws an exception if PageContentType is not Outline. Read-only.|
|parentPage|[Page](page.md)|Gets the page that contains the PageContent object. Read-only.|

## Methods

| Method		   | Return Type	|Description| 
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Deletes the PageContent object.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.|

## Method details

### delete()

Deletes the PageContent object.

#### Syntax

```js
pageContentObject.delete();
```

#### Parameters

None

#### Returns

Void

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
			if(firstPageContent.isNull === false) {
				firstPageContent.delete();
				return context.sync();
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
|param|object|Optional. Accepts parameter and relationship names as a delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns

Void

<br/>

