# Application Object (JavaScript API for Word) ![new](../media/new.jpg) 

Represents the Microsoft Word application.

_Applies to: Word 2016_

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[createDoc(base64File: string)](#createdocbase64file-string)|[Document](document.md)| ![new](../media/new.jpg) Creates a new document by using a base64 encoded .docx file.|

## Method Details


### createDoc(base64File: string)
![new](../media/new.jpg) Creates a new document by using a base64 encoded .docx file.

#### Syntax
```js
applicationObject.createDoc(base64File);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64File|string|Optional. Optional. The base64 encoded .docx file. The default value is null.|

#### Returns
[Document](document.md)
