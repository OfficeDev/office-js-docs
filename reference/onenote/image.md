# Image Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  

Represents an Image. An Image can be a direct child of a PageContent object or a Paragraph object.

To provide feedback on this API, you can [file an issue in GitHub](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image).

## Properties

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|description|string|Gets or sets the description of the Image.|
|height|double|Gets or sets the height of the Image layout.|
|hyperlink|string|Gets or sets the hyperlink of the Image.|
|id|string|Gets the ID of the Image object. Read-only.|
|width|double|Gets or sets the width of the Image layout.|

_See [property access examples](#property-access-examples)_.

## Relationships

| Relationship | Type	|Description| 
|:---------------|:--------|:----------|
|ocrData|[ImageOcrData](imageocrdata.md)|Gets the data obtained by OCR (Optical Character Recognition) of this Image, such as OCR text and language. Read-only.|
|pageContent|[PageContent](pagecontent.md)|Gets the PageContent object that contains the Image. Throws an exception if the Image is not a direct child of a PageContent. This object defines the position of the Image on the page. Read-only.|
|paragraph|[Paragraph](paragraph.md)|Gets the Paragraph object that contains the Image. Throws an exception if the Image is not a direct child of a Paragraph. Read-only.|

## Methods

| Method		   | Return Type	|Description| 
|:---------------|:--------|:----------|
|[getBase64Image()](#getbase64image)|string|Gets the base64-encoded binary representation of the Image.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.|

## Method details

### getBase64Image()

Gets the base64-encoded binary representation of the Image.

#### Syntax

```js
imageObject.getBase64Image();
```

#### Parameters

None

#### Returns

String

#### Examples

```js

var image = null;
var imageString;

OneNote.run(function(ctx){
    // Get the current outline.         
    var outline = ctx.application.getActiveOutline();
    
    // Queue a command to load paragraphs and their types. 
    outline.load("paragraphs/type")
    return ctx.sync().
        then(function(){
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    image = paragraph.image;
                }
            }
        })
        .then(function(){
            if (image != null)
            {
                imageString = image.getBase64Image();
                return ctx.sync();
            }
        })
        .then(function(){
            console.log(imageString);
        });
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

**id, width, height, description and hyperlink**

```js
OneNote.run(function(ctx){
    // Get the current outline.         
    var outline = ctx.application.getActiveOutline();
    var image = null;
    
    // Queue a command to load paragraphs and their types. 
    outline.load("paragraphs/type")
    return ctx.sync().
        then(function(){
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    image = paragraph.image;
                }
            }
        })
        .then(function(){
            if (image != null)
            {
                // load every properties and relationships
                ctx.load(image);
                return ctx.sync();
            }
        })
        .then(function(){
            if (image != null)
            {                   
                console.log("image " + image.id + " width is " + image.width + " height is " + image.height);
                console.log("description: " + image.description);                   
                console.log("hyperlink: " + image.hyperlink);
            }
        });
});
```

<br/>

**ocrData**

```js
var image = null;

OneNote.run(function(ctx){
	// Get the current outline.
	var outline = ctx.application.getActiveOutline();

	// Queue a command to load paragraphs and their types.
	outline.load("paragraphs")
	return ctx.sync().
		then(function(){
			for (var i=0; i < outline.paragraphs.items.length; i++)
			{
				var paragraph = outline.paragraphs.items[i];
				if (paragraph.type == "Image")
				{
					image = paragraph.image;
				}
			}
			if (image != null)
			{
			   image.load("ocrData");
			}
			return ctx.sync();
		})
		.then(function(){
			console.log(image.ocrData);
		});
}).catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>

**paragraph**

```js
OneNote.run(function(ctx){
    // Get the current outline.         
    var outline = ctx.application.getActiveOutline();
    var searchedParagraph = null;
    
    // Queue a command to load paragraphs and their types. 
    outline.load("paragraphs/type")
    return ctx.sync().
        then(function() {
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    searchedParagraph = paragraph;
                    break;
                }
            }
        })
        .then(function() {
            if (searchedParagraph != null)
            {
                // load every properties and relationships
                searchedParagraph.image.load('paragraph');
                return ctx.sync();
            }
        })
        .then(function() {
            if (searchedParagraph != null)
            {                   
                if (searchedParagraph.id != searchedParagraph.image.paragraph.id)
                {
                    console.log("id must match");
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

<br/>

