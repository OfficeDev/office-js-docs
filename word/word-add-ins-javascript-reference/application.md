# Application object (JavaScript API for Word)  ![new](../media/new.jpg) 

Represents the Microsoft Word application. Holds application-level functionalities.

_Applies to: Word 2016, Word for iPad_

## Properties
No properties on this release.

## Relationships
No relationships on this release.

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[createDoc(base64File: string)](#createdoc) ![new](../media/new.jpg)  |[Document](document.md)| Creates a new document, optionally taking an existing one as template.|



## Method details

### createDoc(base64File: string)![new](../media/new.jpg) 

Creates a new document, optionally taking an existing one as template (supplied as a base64 encoded docx).

#### Syntax
```js
var doc = cxt.createDoc();
```

#### Parameters
| Parameter    | Type   |Description|
|:---------------|:--------|:----------|
|base64File|string|Optional. A base64 encoded docx file to be used as template of the newly created document.|

#### Returns
[Document](document.md)

#### Additional details
After a document is created developers can use the existing API to manipulate content on the newly created document. I.e. adding Paragrpahs, content controls, search, header/footer, etc.
USers dont have access to user selection
Document name is automatically created by Word (ie. document1.docx, or the right numeral)
Documetn is created in the same location as the document from where the add-in is executing.


#### Examples
```js
//creates a documents adds a paragrpah at the end of it and finally it opens that document.
Word.run(function (ctx) {
    
    var doc = ctx.application.createDoc();
    ctx.load(doc);

    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
            doc.body.insertParagraph("New Paragraph!", "end");
            doc.open();
            ctx.sync();

    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```



## Support details

Use the [requirement set](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx) in run time checks to make sure your application is supported by the host version of Word. For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx).  