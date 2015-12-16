# Range object (JavaScript API for Word)

Represents a contiguous area in a document.

_Applies to: Word 2016, Word for iPad_

## Properties
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|style|string|Gets or sets the style used for the range. This is the name of the pre-installed or custom style.|
|text|string|Gets the text of the range. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Gets the collection of content control objects that are in the range. Read-only.|
|font|[Font](font.md)|Gets the text format of the range. Use this to get and set font name, size, color, and other properties. Read-only.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Gets the collection of inlinePicture objects that are in the range. Read-only.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Gets the collection of paragraph objects that are in the range. Read-only.|
|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the range. Returns null if there isn't a parent content control. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Clears the contents of the range object. The user can perform the undo operation on the cleared content.|
|[delete()](#delete)|void|Deletes the range and its content from the document.|
|[getHtml()](#gethtml)|string|Gets the HTML representation of the range object.|
|[getOoxml()](#getooxml)|string|Gets the OOXML representation of the range object.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Inserts a break at the specified location. A break can only be inserted into range objects that are contained within the main document body, except if it is a line break which can be inserted into any body object. The insertLocation value can be 'Replace', 'Before' or 'After'.|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Wraps the range object with a rich text content control.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Inserts a document into the range at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts HTML into the range at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Inserts a picture into the range at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts OOXML or wordProcessingML into the range at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserts a paragraph into the range at the specified location. The insertLocation value can be 'Before' or 'After'.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Inserts text into the range at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestringssearchoptions)|[SearchResultCollection](searchresultcollection.md)|Performs a search with the specified searchOptions on the scope of the range object. The search results are a collection of range objects.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Selects and navigates the Word UI to the range. The selectionMode values can be 'Select', 'Start', or 'End'.|
|[adjust(startAdjust: number, endAdjust: number)](#adjust) ![new](../media/new.jpg) | [Range](range.md) | Adjusts the current range to 'startAdjust' positions from start to 'endAdjust' positions to end. Negative values and zero are valid options. |
|[expandTo(range: Range)](#expand) ![new](../media/new.jpg) | [Range](range.md) | Expands the current range until the bounds of the specified Range |
|[getChildRange(rangeOrigin: string, length: number)](#getchildrangerangeorigin-string-length-number) ![new](../media/new.jpg) | [Range](range.md) | Gets a range from the start/end of calling range to the number of specified positions. This method is useful to get the first or last characters within the range | 
|[getRanges(delimiters: string[], excludeDelimiters: bool, trimWhite: bool, excludeEndingMarks: bool, within:bool)](#getrangesdelimiters-string-excludedelimiters-bool-trimwhite-bool-excludeendingmarks-bool)![new](../media/new.jpg) | RangeCollection(rangeCollection.md) | Gets a collection of Ranges within the calling range each one within the specified delimiter(s) and either including or nor the actual delimiters, blanks or ending mark.|
|[hasRange(range: Range, isSubset: bool)](#hasrange) ![new](../media/new.jpg) | bool | Determines whether the calling range contains the specified range. Deverlopers can also determine if its fully contained or not. |


## Method details

### adjust((startAdjust: number, endAdjust: number)  ![new](../media/new.jpg)
Adjusts the current range to 'startAdjust' positions from start to 'endAdjust' positions to end. Negative values and zero are valid options.

#### Parameters
| Parameter    | Type   |Description|
|:---------------|:--------|:----------|
|startAdjust|number|Required. Indicates number of character positions to adjust from the start of the calling range.|
|endAdjust|number|Required. Indicates number of character positios to adjust from the end of the calling range.|

#### Returns
[Range](range.md)

#### Additional details
Negative values will take the positions before start/end
Error conditions:
Method will fail if the start/end adjust values reach start or end of document.



#### Examples
```js

//following example adjusts the current selection to include one character to the left and one to the right.

Word.run(function (ctx) {
    var cc = ctx.document.getSelection();
    ctx.load(cc);

    return ctx.sync().then(function () {
        var r = s.Adjust(-1, 1);  // adjust current range one character to the left and one to the right.
        ctx.load(r);

        return ctx.sync().then(function () {
            console.log("r=[" + r.text + "]");
        }).catch(function(error) {
            console.log(JSON.stringify(error));
        });
    }).catch(function(error) {
        console.log(JSON.stringify(error));
    });
}); 
```

### expandTo(range: Range)  ![new](../media/new.jpg)
Expands the current range until the bounds of the specified Range

#### Parameters
| Parameter    | Type   |Description|
|:---------------|:--------|:----------|
|range |Range|Required. Range to where the calling range will be expanded.|

#### Returns
[Range](range.md)

#### Additional details
Can be used in conjunction with the getChildRange method, for intance, to expand the calling range all the way to the end of the document. Same can be applied to specific paragraph or content control range.


#### Examples
```js


Word.run(function (ctx) {

//following example expands current selection range all the way to the end of the document.
    
    var eof = ctx.document.getChildRange('end', 0); //gets the range at the end of the document.
    ctx.load(cc);

    return ctx.sync().then(function () {
      var currentSelection = ctx.document.getSelection();
      ctx.load(currentSelection);

        return ctx.sync().then(function () {

           var newRange = currentSelection.expand(eof);
           ctx.sync();
        }).catch(function(error) {
            console.log(JSON.stringify(error));
        });
    }).catch(function(error) {
        console.log(JSON.stringify(error));
    });
}); 

```

### getChildRange(rangeOrigin: string, length: number)  ![new](../media/new.jpg)
Gets a range from the start/end of calling range to the number of specified positions. This method is useful to get the first or last characters within the ccalling range. 

#### Parameters
| Parameter    | Type   |Description|
|:---------------|:--------|:----------|
|rangeOrigin|rangeOrigin|Required. Indicates if the Range is to be retrieved from the start or from the end of the calling range. The value can be 'Start' or "End".|
|length|InsertLocation|Required. The number of character positions to be included from the 'Start' or 'End'. 0 is a valid option and will return a range representing the start/end of the calling range|

#### Returns
[Range](range.md)

#### Additional details
This method can be used to get the first or last character of the calling range as Ranges, for instance for formatting purposes. 



#### Examples
```js
// The following example gets the first 10 characters of the current selection.
Word.run(function (ctx) {
    var cc = ctx.document.getSelection();
    ctx.load(cc);

    return ctx.sync().then(function () {
        var r = cc.getChildRange('Start', 10);
        ctx.load(r);

        return ctx.sync().then(function () {
            console.log("r=[" + r.text + "]");
        }).catch(function(error) {
            console.log(JSON.stringify(error));
        });
    }).catch(function(error) {
        console.log(JSON.stringify(error));
    });
}); 
```

### getRanges(delimiters: string[], excludeDelimiters: bool, trimWhite: bool, excludeEndingMarks: bool)  ![new](../media/new.jpg)
Gets a collection of Ranges within the calling range each one within the specified delimiter(s) and either including or nor the actual delimiters, blanks or ending mark.

#### Parameters
| Parameter    | Type   |Description|
|:---------------|:--------|:----------|
|delimiters|string[]|Required. Array of delimiters to be used to populate the resulting collection. Valid  delimiter examplese are " " (space) to get all the words in a given range, "." will get the sentences, etc.|
|excludeDelimiters|bool|Optional. False by Default.  Indicates if the specified delimiters are to be included as individual range objects within the resulting Range collection.|
|trimWhite|bool|Optional. False by Default.  Indicates if spaces, singles spaces or tabs, should be removed from the resulting ranges.|
|excludeEndingMarks|bool|Optional. False by Default.  Indicates if invisible ending marks (such as end of paragraph, end of cell, end of table, line breaks) are to be included as individual ranges within the collection.|
|within|bool|Optional. False by Default.  Indicates if the delimiters will be searched only within the calling range boundaries (true) or if it expands outside of its boundary until delimiters are found top/down.|



#### Returns
[Range](range.md)

#### Additional details
The method will return text (as ranges) within the specified delimiters and if so applicable  within paragraphs, content controls and tables. Text within tables is retrieved cell by cell, left to right (RTL in some cultures) top-down.

When endingMarks option is selected individual ending marks are returned as a single range and coded (i.e End Paragraph mark = '\r'). Supported ending marks: Paragraphs, end of cell and end of row and end of table.

sub documents are not returned, this includes text within shapes, text boxes, comments, etc.

#### Examples
```js
// The following exammple retrieves sentences with different ending delimiters.
Word.run(function (ctx) {
    var delim = [".", "?", ":", "!", "。"];
    var s = ctx.document.body.getRanges(delim, false, true, false, false);
    ctx.load(s);

    return ctx.sync().then(function () {
        for (var i = 0; i < s.items.length; i++)
            console.log("[" + s.items[i].text + "]");
    }).catch(function(error) {
        console.log(JSON.stringify(error));
    });
});
```

### [hasRange(range: Range, isSubset: bool)  ![new](../media/new.jpg)
 Determines whether the calling range contains the specified range. Deverlopers can also determine if its fully contained or not. 

#### Parameters
| Parameter    | Type   |Description|
|:---------------|:--------|:----------|
|range|Range|Required. The Range that will be verified to be contained within the calling range..|
|isSubset|bool|Optional. Used to verify if the supplied range is fully or partially contained within the calling Range.|

#### Returns
Bool

#### Additional details
R1.HasRange(R2, false) returns true if and only if the R1 and R2 overlap.
R1.HasRange(R2, true) returns true if and only if R2 is inside R1.



#### Examples
```js
```

### clear()
Clears the contents of the range object. The user can perform the undo operation on the cleared content.

#### Syntax
```js
rangeObject.clear();
```

#### Parameters
None

#### Returns
void

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to clear the contents of the proxy range object.
    range.clear();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Cleared the selection (range object)');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
### delete()
Deletes the range and its content from the document.

#### Syntax
```js
rangeObject.delete();
```

#### Parameters
None

#### Returns
void

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to delete the range object.
    range.delete();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Deleted the selection (range object)');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### getHtml()
Gets the HTML representation of the range object.

#### Syntax
```js
rangeObject.getHtml();
```

#### Parameters
None

#### Returns
string

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to get the HTML of the current selection. 
    var html = range.getHtml();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The HTML read from the document was: ' + html.value);
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### getOoxml()
Gets the OOXML representation of the range object.

#### Syntax
```js
rangeObject.getOoxml();
```

#### Parameters
None

#### Returns
string

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to get the OOXML of the current selection. 
    var ooxml = range.getOoxml();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The OOXML read from the document was:  ' + ooxml.value);
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Inserts a break at the specified location. A break can only be inserted into range objects that are contained within the main document body, except if it is a line break which can be inserted into any body object. The insertLocation value can be 'Replace', 'Before' or 'After'.

#### Syntax
```js
rangeObject.insertBreak(breakType, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|breakType|BreakType|Required. The break type to add to the range.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Before' or 'After'.|

#### Returns
void

#### Additional details
With the exception of line breaks, you can not insert a break in header, footer, footnote, endnote, comment, and textbox objects. 

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert a page break after the selected text.
    range.insertBreak('page', 'After');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Inserted a page break after the selected text.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertContentControl()
Wraps the range object with a rich text content control.

#### Syntax
```js
rangeObject.insertContentControl();
```

#### Parameters
None

#### Returns
[ContentControl](contentcontrol.md)

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert a content control around the selected text,
    // and create a proxy content control object. We'll update the properties
    // on the content control.
    var myContentControl = range.insertContentControl();
    myContentControl.tag = "Customer-Address";
    myContentControl.title = "Enter Customer Address Here:";
    myContentControl.style = "Normal";
    myContentControl.insertText("One Microsoft Way, Redmond, WA 98052", 'replace');
    myContentControl.cannotEdit = true;
    myContentControl.appearance = "tags";
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Wrapped a content control around the selected text.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Inserts a document into the range at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
rangeObject.insertFileFromBase64(base64File, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64File|string|Required. The file base64 encoded file contents to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert base64 encoded .docx at the beginning of the range.
    // You'll need to implement getBase64() to make this work.
    range.insertFileFromBase64(getBase64(), Word.InsertLocation.start);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Added base64 encoded text to the beginning of the range.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertHtml(html: string, insertLocation: InsertLocation)
Inserts HTML into the range at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
rangeObject.insertHtml(html, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|Required. The HTML to be inserted in the range.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert HTML in to the beginning of the range.
    range.insertHtml('<strong>This is text inserted with range.insertHtml()</strong>', Word.InsertLocation.start);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('HTML added to the beginning of the range.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Inserts a picture into the range at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.

#### Syntax
rangeObject.insertInlinePictureFromBase64(image, insertLocation);

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Required. The base64 encoded image to be inserted in the range.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|

#### Returns
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Inserts OOXML or wordProcessingML into the range at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
rangeObject.insertOoxml(ooxml, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|ooxml|string|Required. The OOXML or wordProcessingML to be inserted in the range.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert OOXML in to the beginning of the range.
    range.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", Word.InsertLocation.start);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('OOXML added to the beginning of the range.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### Additional information
Read [Create better add-ins for Word with Office Open XML](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx) for guidance on working with OOXML.

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Inserts a paragraph into the range at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
rangeObject.insertParagraph(paragraphText, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|paragraphText|string|Required. The paragraph text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Before' or 'After'.|

#### Returns
[Paragraph](paragraph.md)

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert the paragraph after the range.
    range.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Paragraph added to the end of the range.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertText(text: string, insertLocation: InsertLocation)
Inserts text into the range at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
rangeObject.insertText(text, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|text|string|Required. Text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert the paragraph at the end of the range.
    range.insertText('New text inserted into the range.', Word.InsertLocation.end);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Text added to the end of the range.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
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

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to load font and style information for the range.
    context.load(range, 'font/size, font/name, font/color, style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Show the results of the load method. Here we show the
        // property values on the range object.
        var results = "  ---Font size: " + range.font.size +
                      "  ---Font name: " + range.font.name +
                      "  ---Font color: " + range.font.color +
                      "  ---Style: " + range.style;
        console.log(results);
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)
Performs a search with the specified searchOptions on the scope of the range object. The search results are a collection of range objects.

#### Syntax
```js
rangeObject.search(searchText, searchOptions);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|searchText|string|Required. The search text.|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|Optional. Options for the search.|

#### Returns
[SearchResultCollection](searchresultcollection.md)


### select(selectionMode: SelectionMode)
Selects and navigates the Word UI to the range. The selectionMode values can be 'Select', 'Start', or 'End'.

#### Syntax
```js
rangeObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.|

#### Returns
void

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert HTML in to the beginning of the range.
    range.insertHtml('<strong>This is text inserted with range.insertHtml()</strong>', Word.InsertLocation.start);
    
    // Queue a command to select the HTML that was inserted.
    range.select();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Selected the range.');
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