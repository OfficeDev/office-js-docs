# Application Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

The Application object.

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[createDocument(base64File: string)](#createdocumentbase64file-string)|[Document](document.md)|Creates a new document by using a base64 encoded .docx file.|[1.3](../reqset/word-requirement.md)|

## Method Details


### createDocument(base64File: string)
Creates a new document by using a base64 encoded .docx file.

#### Syntax
```js

 Word.run(function (context) {        // lets hold a valid base64 docx on this variable...        
 var myStartingDocAsBase64 = "some valid base64 encoded docx";
 var myNewDoc = context.application.createDocument(myStartingDocAsBase64);  // note that the parameter is optional, a blank doc will be created otherwise               // at this point you can use the entire API on the myNewDoc document.. you can do things like        
 myNewDoc.body.insertParagraph("This is a new paragraph added via API", "end");        //now lets open the document, after this method is called,  you will no longer be able to modify the doc.....        
 myNewDoc.open();
 return context.sync();    
 })
 .catch(function (e) {
 console.log(e.message);
 })

```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|base64File|string|Optional. Optional. The base64 encoded .docx file. The default value is null.|

#### Returns
[Document](document.md)
